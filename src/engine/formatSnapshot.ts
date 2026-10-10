import type { RichAction } from '@/action.types';
import { columnLetterToIndex } from './addressUtils';
import { resolveWorksheet } from './sheetResolve';

/* global Excel */

/**
 * What a formatting action overwrote, read off Excel just before the action
 * runs, so Revert can put it back (TASKS.md #400).
 *
 * The server stores this with the change set and turns it back into
 * FORMAT_RANGE / SET_COLUMN_WIDTH actions (`Server/src/audit/format-snapshot.ts`
 * declares the same shape and checks it; change them together).
 *
 * Only the properties the action writes are read, so a snapshot of a bold and
 * fill change holds bold and fill and nothing else. A snapshot that cannot be
 * taken, or put back exactly, is reported as `restorable: false` with a reason,
 * never skipped: the server refuses a revert it cannot do whole.
 */

export const MAX_SNAPSHOT_CELLS = 20_000;
/** All the formatting of one Accept together. Keeps the apply call's body well under the server's 1 MiB limit. */
export const MAX_SNAPSHOT_TOTAL_CELLS = 60_000;
/** Distinct formats in one range. A range where every cell differs would send a format per cell. */
export const MAX_SNAPSHOT_PALETTE = 2_000;
export const MAX_SNAPSHOT_COLUMNS = 200;
/** The last cell of a worksheet (XFD1048576), 0-based. */
const LAST_ROW = 1_048_575;
const LAST_COLUMN = 16_383;

/** Cells still available to the snapshots of one Accept; shared across its actions and counted down. */
export interface SnapshotBudget {
  cells: number;
}

export interface SnapshotCellFormat {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontName?: string;
  fontColor?: string;
  fillColor?: string;
  clearFill?: boolean;
  numberFormat?: string;
  horizontalAlignment?: 'left' | 'center' | 'right' | 'general';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
}

export interface FormatSnapshot {
  kind: 'format' | 'columns';
  sheetName: string;
  restorable: boolean;
  reason?: string;
  row?: number;
  col?: number;
  rowCount?: number;
  colCount?: number;
  palette?: SnapshotCellFormat[];
  grid?: number[][];
  widths?: { col: number; width: number }[];
}

type FormatAction = Extract<RichAction, { type: 'FORMAT_RANGE' }>;
type AutofitAction = Extract<RichAction, { type: 'AUTOFIT_COLUMNS' }>;

function refusal(kind: FormatSnapshot['kind'], sheetName: string, reason: string): FormatSnapshot {
  return { kind, sheetName, restorable: false, reason };
}

const H_ALIGN: Record<string, SnapshotCellFormat['horizontalAlignment']> = {
  General: 'general',
  Left: 'left',
  Center: 'center',
  Right: 'right',
};
const V_ALIGN: Record<string, SnapshotCellFormat['verticalAlignment']> = {
  Top: 'top',
  Center: 'middle',
  Bottom: 'bottom',
};

/** Thrown inside the reader when a cell holds something FORMAT_RANGE cannot write back. */
class NotRestorable extends Error {}

interface Touched {
  bold: boolean;
  italic: boolean;
  fontSize: boolean;
  fontName: boolean;
  fontColor: boolean;
  fill: boolean;
  numberFormat: boolean;
  horizontalAlignment: boolean;
  verticalAlignment: boolean;
  wrapText: boolean;
}

function touchedBy(format: FormatAction['format']): Touched {
  const f = (format ?? {}) as Record<string, unknown>;
  const has = (key: string) => f[key] !== undefined;
  return {
    bold: has('bold'),
    italic: has('italic'),
    fontSize: has('fontSize'),
    fontName: has('fontName'),
    fontColor: has('fontColor'),
    fill: has('fillColor') || f.clearFill === true,
    numberFormat: has('numberFormat'),
    horizontalAlignment: has('horizontalAlignment'),
    verticalAlignment: has('verticalAlignment'),
    wrapText: has('wrapText'),
  };
}

function loadOptions(touched: Touched): Excel.CellPropertiesLoadOptions {
  const font: Excel.CellPropertiesFontLoadOptions = {};
  if (touched.bold) font.bold = true;
  if (touched.italic) font.italic = true;
  if (touched.fontSize) font.size = true;
  if (touched.fontName) font.name = true;
  if (touched.fontColor) font.color = true;
  const format: Excel.CellPropertiesFormatLoadOptions = {};
  if (Object.keys(font).length > 0) format.font = font;
  if (touched.fill) format.fill = { ...FILL_LOAD };
  if (touched.horizontalAlignment) format.horizontalAlignment = true;
  if (touched.verticalAlignment) format.verticalAlignment = true;
  if (touched.wrapText) format.wrapText = true;
  return { format };
}

function need<T>(value: T | null | undefined, what: string): T {
  if (value === null || value === undefined) throw new NotRestorable(`${what} could not be read`);
  return value;
}

type FillState = { clearFill: true } | { fillColor: string };

/** A fill as Excel reports it for one cell. */
interface RawFill {
  color?: unknown;
  pattern?: unknown;
  tintAndShade?: unknown;
}

const FILL_LOAD = { color: true, pattern: true, tintAndShade: true } as const;

const blankOrWhite = (color: unknown) => typeof color === 'string' && ['', '#FFFFFF'].includes(color.trim().toUpperCase());
const sameFill = (a: RawFill | undefined, b: RawFill | undefined) =>
  String(a?.color ?? '').trim().toUpperCase() === String(b?.color ?? '').trim().toUpperCase() &&
  (a?.pattern ?? null) === (b?.pattern ?? null) &&
  (a?.tintAndShade ?? null) === (b?.tintAndShade ?? null);

/**
 * One cell's fill as "none" or a solid colour.
 *
 * Whether a cell has a fill at all is its pattern, not its colour, and where
 * Excel reports the pattern that settles it. Excel on Windows did not
 * (Oct 9 2026, live): a solid light-green cell read as
 * `{ color: "#C6EFCE", pattern: null }`, for the cell and for the range alike.
 * Without a pattern only the colour is left, and the colour alone cannot tell
 * "no fill" from a white fill. So the cell is compared with `reference`, a cell
 * at the far corner of the sheet that nobody formats: a cell that reads exactly
 * like it has no fill.
 *
 * What this cannot do: where Excel reports an explicit white fill exactly as it
 * reports no fill, the two are the same to every read there is, and a white
 * fill is put back as no fill. Any other colour is exact.
 */
function fillFrom(cell: RawFill | undefined, reference: RawFill | undefined): FillState {
  const pattern = cell?.pattern;
  if (pattern === 'None') return { clearFill: true };
  if (typeof pattern === 'string' && pattern && pattern !== 'Solid') {
    throw new NotRestorable(`fill pattern "${pattern}" cannot be put back`);
  }
  if (typeof cell?.color !== 'string') throw new NotRestorable('fill colour could not be read');
  const color = cell.color.trim();
  if (pattern === 'Solid') {
    if (!color) throw new NotRestorable('fill colour could not be read');
    return { fillColor: color };
  }

  // No pattern reported. The corner cell counts as "no fill" only if it looks like it:
  // on a sheet filled edge to edge it has that fill too, and matching it means nothing.
  const referenceIsUnfilled =
    blankOrWhite(reference?.color) && !(typeof reference?.pattern === 'string' && reference.pattern !== 'None');
  if (referenceIsUnfilled && sameFill(cell, reference)) return { clearFill: true };
  if (color === '') return { clearFill: true };
  if (!blankOrWhite(color)) return { fillColor: color };
  // White, and not the same as an unfilled cell: an explicit white fill, if there is an unfilled cell to compare with.
  if (referenceIsUnfilled) return { fillColor: color };
  throw new NotRestorable('a white fill could not be told from no fill on this sheet');
}

function cellFormat(
  touched: Touched,
  props: Excel.CellProperties | undefined,
  numberFormat: unknown,
  fill: FillState | undefined,
): SnapshotCellFormat {
  const out: SnapshotCellFormat = {};
  const font = props?.format?.font;
  if (touched.bold) out.bold = need(font?.bold, 'bold');
  if (touched.italic) out.italic = need(font?.italic, 'italic');
  if (touched.fontSize) out.fontSize = need(font?.size, 'font size');
  if (touched.fontName) out.fontName = need(font?.name, 'font name');
  if (touched.fontColor) {
    // Excel reports "no explicit colour" as an empty string; black is what it draws.
    const color = need(font?.color, 'font colour');
    out.fontColor = color.trim() ? color : '#000000';
  }
  if (touched.fill) Object.assign(out, need(fill, 'fill'));
  if (touched.numberFormat) {
    if (typeof numberFormat !== 'string') throw new NotRestorable('number format could not be read');
    out.numberFormat = numberFormat;
  }
  if (touched.horizontalAlignment) {
    const value = H_ALIGN[String(need(props?.format?.horizontalAlignment, 'alignment'))];
    if (!value) throw new NotRestorable('this horizontal alignment cannot be put back');
    out.horizontalAlignment = value;
  }
  if (touched.verticalAlignment) {
    const value = V_ALIGN[String(need(props?.format?.verticalAlignment, 'vertical alignment'))];
    if (!value) throw new NotRestorable('this vertical alignment cannot be put back');
    out.verticalAlignment = value;
  }
  if (touched.wrapText) out.wrapText = need(props?.format?.wrapText, 'wrap text');
  return out;
}

async function captureRangeFormat(
  action: FormatAction,
  ctx: Excel.RequestContext,
  budget: SnapshotBudget,
): Promise<FormatSnapshot> {
  const sheetName = action.sheetName;
  if ((action.format as { borders?: unknown } | undefined)?.borders) {
    return refusal('format', sheetName, 'borders cannot be put back edge by edge');
  }
  const touched = touchedBy(action.format);
  const sheet = resolveWorksheet(ctx, sheetName);
  const range = sheet.getRange(action.range);
  range.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
  await ctx.sync();

  const { rowIndex, columnIndex, rowCount, columnCount } = range;
  const cells = rowCount * columnCount;
  if (cells > MAX_SNAPSHOT_CELLS) {
    return refusal('format', sheetName, `range has more than ${MAX_SNAPSHOT_CELLS} cells`);
  }
  if (cells > budget.cells) {
    return refusal('format', sheetName, `this change formats more than ${MAX_SNAPSHOT_TOTAL_CELLS} cells in all`);
  }
  if (!Object.values(touched).some(Boolean)) {
    return { kind: 'format', sheetName, restorable: true, row: rowIndex, col: columnIndex, rowCount: 0, colCount: 0, palette: [], grid: [] };
  }

  budget.cells -= cells;
  const props = range.getCellProperties(loadOptions(touched));
  if (touched.numberFormat) range.load('numberFormat');
  // The sheet's last cell: what a cell with no fill reads as in this Excel. See fillFrom.
  const referenceProps = touched.fill
    ? sheet.getRangeByIndexes(LAST_ROW, LAST_COLUMN, 1, 1).getCellProperties({ format: { fill: { ...FILL_LOAD } } })
    : undefined;
  await ctx.sync();

  try {
    // Excel throws on reading a property that was not loaded, so each is read only when it was.
    const numberFormats = touched.numberFormat ? range.numberFormat : undefined;
    const reference = referenceProps?.value?.[0]?.[0]?.format?.fill as RawFill | undefined;
    const palette: SnapshotCellFormat[] = [];
    const seen = new Map<string, number>();
    const grid: number[][] = [];
    for (let r = 0; r < rowCount; r += 1) {
      const line: number[] = [];
      for (let c = 0; c < columnCount; c += 1) {
        const cellProps = props.value?.[r]?.[c];
        const fill = touched.fill ? fillFrom(cellProps?.format?.fill as RawFill | undefined, reference) : undefined;
        const cell = cellFormat(touched, cellProps, numberFormats?.[r]?.[c], fill);
        const key = JSON.stringify(cell);
        let index = seen.get(key);
        if (index === undefined) {
          if (palette.length >= MAX_SNAPSHOT_PALETTE) {
            throw new NotRestorable(`more than ${MAX_SNAPSHOT_PALETTE} different formats in one range`);
          }
          index = palette.length;
          palette.push(cell);
          seen.set(key, index);
        }
        line.push(index);
      }
      grid.push(line);
    }
    return { kind: 'format', sheetName, restorable: true, row: rowIndex, col: columnIndex, rowCount, colCount: columnCount, palette, grid };
  } catch (error) {
    // What Excel actually returned, for the one cell that is enough to see the shape.
    // Formatting only, never cell contents. This is what turned a guess into a fix.
    console.warn(
      '[Cellix] Format snapshot not taken:',
      error instanceof Error ? error.message : String(error),
      JSON.stringify({
        range: action.range,
        firstCell: props.value?.[0]?.[0]?.format,
        unfilledReference: referenceProps?.value?.[0]?.[0]?.format?.fill,
      }).slice(0, 600),
    );
    throw error;
  }
}

async function captureColumnWidths(action: AutofitAction, ctx: Excel.RequestContext): Promise<FormatSnapshot> {
  const sheetName = action.sheetName;
  const sheet = resolveWorksheet(ctx, sheetName);

  let indices: number[];
  if (action.columns && action.columns.length > 0) {
    if (action.columns.some((c) => !/^[A-Za-z]{1,3}$/.test(String(c)))) {
      return refusal('columns', sheetName, 'a column name was not recognised');
    }
    indices = action.columns.map((c) => columnLetterToIndex(String(c)));
  } else {
    const used = sheet.getUsedRangeOrNullObject();
    used.load(['isNullObject', 'columnIndex', 'columnCount']);
    await ctx.sync();
    indices = used.isNullObject ? [] : Array.from({ length: used.columnCount }, (_, i) => used.columnIndex + i);
  }
  if (indices.length > MAX_SNAPSHOT_COLUMNS) {
    return refusal('columns', sheetName, `more than ${MAX_SNAPSHOT_COLUMNS} columns`);
  }

  const formats = indices.map((index) => sheet.getRangeByIndexes(0, index, 1, 1).format);
  formats.forEach((format) => format.load('columnWidth'));
  await ctx.sync();
  return {
    kind: 'columns',
    sheetName,
    restorable: true,
    widths: indices.map((col, i) => ({ col, width: formats[i].columnWidth })),
  };
}

/**
 * The snapshot for a formatting action, or null for any other action. Never
 * throws: a failure to read becomes a snapshot that refuses the revert.
 */
export async function captureFormatSnapshot(
  action: RichAction,
  ctx: Excel.RequestContext,
  budget: SnapshotBudget = { cells: MAX_SNAPSHOT_TOTAL_CELLS },
): Promise<FormatSnapshot | null> {
  if (action.type !== 'FORMAT_RANGE' && action.type !== 'AUTOFIT_COLUMNS') return null;
  const kind = action.type === 'FORMAT_RANGE' ? 'format' : 'columns';
  try {
    return action.type === 'FORMAT_RANGE'
      ? await captureRangeFormat(action, ctx, budget)
      : await captureColumnWidths(action, ctx);
  } catch (error) {
    const reason =
      error instanceof NotRestorable
        ? error.message
        : `previous formatting could not be read: ${error instanceof Error ? error.message : String(error)}`;
    return refusal(kind, action.sheetName, reason.slice(0, 200));
  }
}
