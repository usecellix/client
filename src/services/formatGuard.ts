import { FormatSpec, SheetAction } from '@/types/sheet-actions';

/* global Excel */

type FormatSource = 'above' | 'below' | 'left' | { address: string };

export interface BorderSide {
  style: string;
  color: string;
}

export interface BorderSet {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
}

export interface CellFormat {
  numberFormat?: string;
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
  };
  fill?: {
    color?: string;
  };
  borders?: BorderSet;
  // Index off RangeFormat: Office.js widens these to the enum plus its string
  // aliases, and we round-trip values read straight back off a range.
  horizontalAlignment?: Excel.RangeFormat['horizontalAlignment'];
  verticalAlignment?: Excel.RangeFormat['verticalAlignment'];
  wrapText?: boolean;
}

const DATE_FORMAT_MAP: Record<string, string> = {
  'DD/MM/YYYY': 'd/m/yyyy',
  'MM/DD/YYYY': 'm/d/yyyy',
  'YYYY-MM-DD': 'yyyy-mm-dd',
  'DD-MMM-YY': 'd-mmm-yy',
  'DD-MMM-YYYY': 'd-mmm-yyyy',
  'MMMM D, YYYY': 'mmmm d, yyyy',
  'DD/MM/YY': 'd/m/yy',
  'M/D/YYYY h:mm': 'm/d/yyyy h:mm',
};

function isBlankCell(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function isDateNumberFormat(numberFormat?: string): boolean {
  if (!numberFormat) return false;
  return detectDateFormat(numberFormat) !== null;
}

/** True when Excel will treat the cell as unformatted / General. */
export function isGeneralOrEmptyFormat(numberFormat?: string): boolean {
  if (!numberFormat) return true;
  const lower = normalizeNumberFormat(numberFormat).toLowerCase();
  return !lower || lower === 'general';
}

/**
 * Prefer an explicit format, else the cell's current format, else a
 * non-General format from the column cell above (sheet convention).
 */
export function resolvePreservedNumberFormat(
  existingFormat?: string,
  aboveFormat?: string,
  explicitFormat?: string,
): string {
  if (explicitFormat && !isGeneralOrEmptyFormat(explicitFormat)) {
    return explicitFormat;
  }
  if (existingFormat && !isGeneralOrEmptyFormat(existingFormat)) {
    return existingFormat;
  }
  if (aboveFormat && !isGeneralOrEmptyFormat(aboveFormat)) {
    return aboveFormat;
  }
  return existingFormat ?? explicitFormat ?? aboveFormat ?? 'General';
}

/** Coerce date-like text to Date/serial when the target format is a date format. */
export function coerceValueForNumberFormat(
  value: unknown,
  numberFormat?: string,
): unknown {
  if (!isDateNumberFormat(numberFormat)) return value ?? '';
  return parseDateValue(value, numberFormat);
}

function isMonthFirstDateFormat(numberFormat: string): boolean {
  const lower = normalizeNumberFormat(numberFormat).toLowerCase();
  const monthIndex = lower.search(/m+/);
  const dayIndex = lower.search(/d+/);
  if (monthIndex < 0 || dayIndex < 0) return false;
  return monthIndex < dayIndex;
}

function parseDateValue(value: unknown, numberFormat?: string): unknown {
  if (value instanceof Date) return value;
  if (typeof value === 'number') {
    if (value > 25000 && value < 80000) return value;
    return value;
  }
  if (isBlankCell(value)) return value;

  const text = String(value).trim();

  const namedMonth = text.match(/^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{2,4})$/);
  if (namedMonth) {
    const parsed = new Date(`${namedMonth[2]} ${namedMonth[1]}, ${namedMonth[3]}`);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const parsed = new Date(`${text}T00:00:00`);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  const slashParts = text.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
  if (slashParts) {
    const first = Number(slashParts[1]);
    const second = Number(slashParts[2]);
    let year = Number(slashParts[3]);
    if (year < 100) year += 2000;

    const monthFirst = numberFormat ? isMonthFirstDateFormat(numberFormat) : false;
    const day = monthFirst ? second : first;
    const month = monthFirst ? first : second;
    const parsed = new Date(year, month - 1, day);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  const fallback = Date.parse(text);
  if (!Number.isNaN(fallback)) return new Date(fallback);

  return value;
}

function getReferenceRowIndex(
  source: FormatSource,
  row: number,
  rowCount: number,
): number | null {
  if (source === 'above') return row > 0 ? row - 1 : null;
  if (source === 'below') return row + rowCount;
  return null;
}

async function readReferenceNumberFormats(
  context: Excel.RequestContext,
  worksheet: Excel.Worksheet,
  refRow: number,
  startCol: number,
  colCount: number,
): Promise<string[]> {
  const refRange = worksheet.getRangeByIndexes(refRow, startCol, 1, colCount);
  refRange.load('numberFormat');
  await context.sync();

  const formats = refRange.numberFormat as string[][];
  return Array.from({ length: colCount }, (_, index) => String(formats[0]?.[index] ?? ''));
}

export async function coerceRowDataToReferenceFormats(
  context: Excel.RequestContext,
  worksheet: Excel.Worksheet,
  targetRow: number,
  source: FormatSource,
  data: unknown[],
  startCol: number,
  colCount: number,
): Promise<unknown[]> {
  const refRow = getReferenceRowIndex(source, targetRow, 1);
  if (refRow == null || refRow < 0) {
    return Array.from({ length: colCount }, (_, index) => data[index] ?? '');
  }

  const formats = await readReferenceNumberFormats(
    context,
    worksheet,
    refRow,
    startCol,
    colCount,
  );

  return Array.from({ length: colCount }, (_, index) => {
    const value = index < data.length ? data[index] : '';
    const format = formats[index] ?? '';
    return isDateNumberFormat(format) ? parseDateValue(value, format) : (value ?? '');
  });
}

async function copyReferenceFormats(
  worksheet: Excel.Worksheet,
  source: FormatSource,
  row: number,
  col: number,
  rowCount: number,
  colCount: number,
): Promise<boolean> {
  const refRow = getReferenceRowIndex(source, row, rowCount);
  if (refRow == null || refRow < 0) return false;

  const spanRowCount = row - refRow + rowCount;
  if (spanRowCount < 1) return false;

  const refRange = worksheet.getRangeByIndexes(refRow, col, 1, colCount);
  const fillRange = worksheet.getRangeByIndexes(refRow, col, spanRowCount, colCount);
  refRange.autoFill(fillRange, Excel.AutoFillType.fillFormats);
  return true;
}

function mergeFormatsForRowAppend(reference: CellFormat, explicit?: CellFormat): CellFormat {
  if (!explicit) return reference;

  const merged = mergeFormats(reference, explicit);
  return {
    ...merged,
    numberFormat: reference.numberFormat,
    horizontalAlignment: reference.horizontalAlignment ?? merged.horizontalAlignment,
    verticalAlignment: reference.verticalAlignment ?? merged.verticalAlignment,
  };
}

function formatSpecToCellFormat(spec?: FormatSpec): CellFormat | undefined {
  if (!spec) return undefined;

  const borders: BorderSet | undefined =
    spec.borders && spec.borders !== 'none'
      ? {
          top: { style: 'Continuous', color: '#000000' },
          bottom: { style: 'Continuous', color: '#000000' },
          left: { style: 'Continuous', color: '#000000' },
          right: { style: 'Continuous', color: '#000000' },
        }
      : undefined;

  const hAlignMap: Record<string, Excel.RangeFormat['horizontalAlignment']> = {
    left: 'Left',
    center: 'Center',
    right: 'Right',
  };

  const vAlignMap: Record<string, Excel.RangeFormat['verticalAlignment']> = {
    top: 'Top',
    middle: 'Center',
    bottom: 'Bottom',
  };

  return {
    numberFormat: spec.numberFormat,
    font: {
      size: spec.fontSize,
      bold: spec.bold,
      italic: spec.italic,
      color: spec.fontColor,
    },
    fill: spec.fillColor ? { color: spec.fillColor } : undefined,
    borders,
    horizontalAlignment: spec.horizontalAlignment
      ? hAlignMap[spec.horizontalAlignment]
      : undefined,
    verticalAlignment: spec.verticalAlignment
      ? vAlignMap[spec.verticalAlignment]
      : undefined,
    wrapText: spec.wrapText,
  };
}

export async function readReferenceFormat(
  context: Excel.RequestContext,
  worksheet: Excel.Worksheet,
  source: FormatSource,
  targetRow: number,
  targetCol: number,
  rowCount = 1,
  colCount = 1,
): Promise<CellFormat> {
  let refRange: Excel.Range;

  if (source === 'above') {
    if (targetRow <= 0) throw new Error('No reference row above');
    refRange = worksheet.getRangeByIndexes(targetRow - 1, targetCol, 1, colCount);
  } else if (source === 'below') {
    refRange = worksheet.getRangeByIndexes(targetRow + rowCount, targetCol, 1, colCount);
  } else if (source === 'left') {
    if (targetCol <= 0) throw new Error('No reference column to the left');
    refRange = worksheet.getRangeByIndexes(targetRow, targetCol - 1, rowCount, 1);
  } else {
    refRange = worksheet.getRange(source.address);
  }

  const refRow =
    source === 'above'
      ? targetRow - 1
      : source === 'below'
        ? targetRow + rowCount
        : source === 'left'
          ? targetRow
          : targetRow;
  const refCol =
    source === 'left' ? targetCol - 1 : targetCol;

  const firstCell =
    typeof source === 'object'
      ? refRange
      : worksheet.getRangeByIndexes(refRow, refCol, 1, 1);

  firstCell.load([
    'numberFormat',
    'format/fill/color',
    'format/font/name',
    'format/font/size',
    'format/font/bold',
    'format/font/italic',
    'format/font/color',
    'format/horizontalAlignment',
    'format/verticalAlignment',
    'format/wrapText',
    'format/borders/top/style',
    'format/borders/top/color',
    'format/borders/bottom/style',
    'format/borders/bottom/color',
    'format/borders/left/style',
    'format/borders/left/color',
    'format/borders/right/style',
    'format/borders/right/color',
  ]);

  await context.sync();

  const fmt = firstCell.format;
  const numberFormat = Array.isArray(firstCell.numberFormat)
    ? String(firstCell.numberFormat[0]?.[0] ?? '')
    : '';

  const borders: BorderSet = {
    top: {
      style: fmt.borders.getItem('EdgeTop').style,
      color: fmt.borders.getItem('EdgeTop').color,
    },
    bottom: {
      style: fmt.borders.getItem('EdgeBottom').style,
      color: fmt.borders.getItem('EdgeBottom').color,
    },
    left: {
      style: fmt.borders.getItem('EdgeLeft').style,
      color: fmt.borders.getItem('EdgeLeft').color,
    },
    right: {
      style: fmt.borders.getItem('EdgeRight').style,
      color: fmt.borders.getItem('EdgeRight').color,
    },
  };

  return {
    numberFormat: numberFormat || undefined,
    font: {
      name: fmt.font.name,
      size: fmt.font.size,
      bold: fmt.font.bold,
      italic: fmt.font.italic,
      color: fmt.font.color,
    },
    fill: {
      color: fmt.fill.color ?? undefined,
    },
    borders,
    horizontalAlignment: fmt.horizontalAlignment,
    verticalAlignment: fmt.verticalAlignment,
    wrapText: fmt.wrapText,
  };
}

export function mergeFormats(reference: CellFormat, explicit?: CellFormat): CellFormat {
  if (!explicit) return reference;

  return {
    numberFormat: explicit.numberFormat ?? reference.numberFormat,
    font: {
      name: explicit.font?.name ?? reference.font?.name,
      size: explicit.font?.size ?? reference.font?.size,
      bold: explicit.font?.bold ?? reference.font?.bold,
      italic: explicit.font?.italic ?? reference.font?.italic,
      color: explicit.font?.color ?? reference.font?.color,
    },
    fill: {
      color: explicit.fill?.color ?? reference.fill?.color,
    },
    borders: explicit.borders ?? reference.borders,
    horizontalAlignment: explicit.horizontalAlignment ?? reference.horizontalAlignment,
    verticalAlignment: explicit.verticalAlignment ?? reference.verticalAlignment,
    wrapText: explicit.wrapText ?? reference.wrapText,
  };
}

export async function applyFormat(
  range: Excel.Range,
  format: CellFormat,
): Promise<void> {
  if (format.numberFormat != null) {
    range.numberFormat = [[format.numberFormat]];
  }

  const fmt = range.format;

  if (format.font) {
    if (format.font.name != null) fmt.font.name = format.font.name;
    if (format.font.size != null) fmt.font.size = format.font.size;
    if (format.font.bold != null) fmt.font.bold = format.font.bold;
    if (format.font.italic != null) fmt.font.italic = format.font.italic;
    if (format.font.color != null) fmt.font.color = format.font.color;
  }

  if (format.fill?.color) {
    fmt.fill.pattern = 'Solid';
    fmt.fill.color = format.fill.color;
  }

  if (format.borders) {
    // Office.js names these BorderIndex / BorderLineStyle — BorderSide and
    // BorderStyle do not exist in the Excel namespace.
    type BorderEdge = Parameters<Excel.RangeBorderCollection['getItem']>[0];
    const applyBorder = (side: keyof BorderSet, excelSide: BorderEdge) => {
      const b = format.borders![side];
      if (!b) return;
      const border = fmt.borders.getItem(excelSide);
      if (b.style && b.style !== 'None') {
        border.style = b.style as Excel.RangeBorder['style'];
        border.color = b.color ?? '#000000';
      }
    };
    applyBorder('top', 'EdgeTop');
    applyBorder('bottom', 'EdgeBottom');
    applyBorder('left', 'EdgeLeft');
    applyBorder('right', 'EdgeRight');
  }

  if (format.horizontalAlignment != null) fmt.horizontalAlignment = format.horizontalAlignment;
  if (format.verticalAlignment != null) fmt.verticalAlignment = format.verticalAlignment;
  if (format.wrapText != null) fmt.wrapText = format.wrapText;
}

function resolveFormatSource(action: SheetAction): FormatSource | null {
  if (action.type === 'ADD_ROW') return 'above';
  if (action.type === 'INSERT_ROW') {
    return action.position === 'above' ? 'below' : 'above';
  }
  if (action.type === 'SET_CELL' || action.type === 'SET_FORMULA') {
    return 'above';
  }
  return null;
}

/**
 * Inherit adjacent-row format and merge with any explicit action.format.
 * Call AFTER writing cell values. Does not call context.sync().
 */
export async function applyFormatGuard(
  context: Excel.RequestContext,
  worksheet: Excel.Worksheet,
  action: SheetAction,
  row: number,
  col: number,
  rowCount: number,
  colCount: number,
): Promise<void> {
  const needsInheritance = ['ADD_ROW', 'INSERT_ROW', 'SET_CELL', 'SET_FORMULA'].includes(
    action.type,
  );
  if (!needsInheritance) return;

  const source = resolveFormatSource(action);
  if (!source) return;

  const explicit = formatSpecToCellFormat(action.format);

  // SET_CELL / SET_FORMULA always re-apply column format after write so Excel
  // does not replace sheet date formats with a locale default.

  if (action.type === 'ADD_ROW' || action.type === 'INSERT_ROW') {
    try {
      await copyReferenceFormats(worksheet, source, row, col, rowCount, colCount);
    } catch {
      // Per-column inheritance below still runs.
    }
  }

  const isRowAppend = action.type === 'ADD_ROW' || action.type === 'INSERT_ROW';

  for (let columnOffset = 0; columnOffset < colCount; columnOffset += 1) {
    const cellRange = worksheet.getRangeByIndexes(row, col + columnOffset, rowCount, 1);
    const refCol = col + columnOffset;

    try {
      const refFormat = await readReferenceFormat(
        context,
        worksheet,
        source,
        row,
        refCol,
        rowCount,
        1,
      );
      const merged = isRowAppend
        ? mergeFormatsForRowAppend(refFormat, explicit)
        : mergeFormats(refFormat, explicit);

      if (action.type === 'SET_CELL' || action.type === 'ADD_ROW' || action.type === 'INSERT_ROW') {
        cellRange.load('values');
        await context.sync();
        const currentValue = cellRange.values?.[0]?.[0];
        if (!isBlankCell(currentValue) && isDateNumberFormat(merged.numberFormat)) {
          const coerced = parseDateValue(currentValue, merged.numberFormat);
          if (coerced !== currentValue) {
            cellRange.values = [[coerced]];
          }
        }
      }

      await applyFormat(cellRange, merged);
    } catch {
      if (explicit) {
        await applyFormat(cellRange, explicit);
      }
    }
  }
}

export function detectDateFormat(numberFormat: string): string | null {
  if (!numberFormat) return null;

  const normalized = normalizeNumberFormat(numberFormat);
  const lower = normalized.toLowerCase();

  if (!lower.includes('y') && !lower.includes('m') && !lower.includes('d')) {
    return null;
  }

  for (const [canonical, excelFmt] of Object.entries(DATE_FORMAT_MAP)) {
    if (lower === excelFmt.toLowerCase()) return canonical;
  }

  return normalized;
}

function normalizeNumberFormat(numberFormat: string): string {
  return numberFormat.split(';')[0]?.replace(/\[[^\]]*\]/g, '').trim() ?? '';
}

/**
 * Write values while keeping each cell's sheet numberFormat (especially dates).
 * Call AFTER obtaining `range`; does not call context.sync() after the write
 * (caller should sync). Loads + syncs once (or twice if inheriting from above).
 */
export async function writeRangeValuesPreservingNumberFormat(
  context: Excel.RequestContext,
  worksheet: Excel.Worksheet,
  range: Excel.Range,
  values: unknown[][],
  options?: { explicitNumberFormat?: string },
): Promise<void> {
  range.load(['numberFormat', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
  await context.sync();

  const rowCount = range.rowCount;
  const colCount = range.columnCount;
  const existing = (range.numberFormat as string[][]) ?? [];

  let aboveFormats: string[] | null = null;
  if (range.rowIndex > 0) {
    const above = worksheet.getRangeByIndexes(
      range.rowIndex - 1,
      range.columnIndex,
      1,
      colCount,
    );
    above.load('numberFormat');
    await context.sync();
    const aboveMatrix = (above.numberFormat as string[][]) ?? [];
    aboveFormats = Array.from({ length: colCount }, (_, c) =>
      String(aboveMatrix[0]?.[c] ?? ''),
    );
  }

  const formats: string[][] = [];
  const coerced: (string | number | boolean | Date)[][] = [];

  for (let r = 0; r < rowCount; r += 1) {
    formats[r] = [];
    coerced[r] = [];
    for (let c = 0; c < colCount; c += 1) {
      const fmt = resolvePreservedNumberFormat(
        String(existing[r]?.[c] ?? ''),
        aboveFormats?.[c],
        options?.explicitNumberFormat,
      );
      formats[r]![c] = fmt;
      const raw = values[r]?.[c];
      const next = coerceValueForNumberFormat(raw, fmt);
      if (
        next === null ||
        next === undefined ||
        typeof next === 'string' ||
        typeof next === 'number' ||
        typeof next === 'boolean' ||
        next instanceof Date
      ) {
        coerced[r]![c] = (next ?? '') as string | number | boolean | Date;
      } else {
        coerced[r]![c] = String(next);
      }
    }
  }

  range.values = coerced as (string | number | boolean)[][];
  range.numberFormat = formats;
}

/**
 * Snapshot numberFormat, run a write, then restore formats (optionally after
 * permuting the format matrix the same way as values — e.g. sort).
 */
export async function preserveNumberFormatsAroundWrite(
  range: Excel.Range,
  context: Excel.RequestContext,
  write: () => void | Promise<void>,
  remapFormats?: (formats: string[][]) => string[][],
): Promise<void> {
  range.load('numberFormat');
  await context.sync();
  const snapshot = ((range.numberFormat as string[][]) ?? []).map((row) =>
    row.map((cell) => String(cell ?? '')),
  );
  await write();
  range.numberFormat = remapFormats ? remapFormats(snapshot) : snapshot;
}
