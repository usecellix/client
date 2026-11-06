import { ColumnMeta, SheetSnapshot, WorkbookContext } from '@/types/cellix.types';
import { detectDateFormat } from '@/services/formatGuard';
import { DeepWorkbookContext, SheetContext } from './context.types';

function isBlankCell(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function detectColumnType(
  nonEmpty: unknown[],
  fmt: string,
): ColumnMeta['detectedType'] {
  const dateFormat = detectDateFormat(fmt);
  if (dateFormat) return 'date';
  if (nonEmpty.every((v) => typeof v === 'number')) return 'number';
  if (nonEmpty.every((v) => typeof v === 'boolean')) return 'boolean';
  if (fmt.includes('$') || fmt.includes('₹') || fmt.includes('€')) return 'currency';
  if (nonEmpty.length > 0) return 'text';
  return 'unknown';
}

/** Data rows start immediately after the header row — not always index 1. */
function buildHeadTailSample(
  values: unknown[][],
  headerRowIndex: number,
  head = 5,
  tail = 5,
): (string | number | null)[][] {
  const dataStart = headerRowIndex + 1;
  if (values.length <= dataStart) return [];
  const dataRows = values.slice(dataStart);
  const selected =
    dataRows.length <= head + tail
      ? dataRows
      : [...dataRows.slice(0, head), ...dataRows.slice(-tail)];

  return selected.map((row) =>
    row.map((v) => (v === '' || v == null ? null : (v as string | number))),
  );
}

/** Distinct non-General number formats in a column's data cells, most common first (≤8). */
export function countColumnFormats(
  numberFormats: string[][],
  headerRowIndex: number,
  colIdx: number,
): Array<{ format: string; count: number }> {
  const counts = new Map<string, number>();
  for (let r = headerRowIndex + 1; r < numberFormats.length; r += 1) {
    const f = String(numberFormats[r]?.[colIdx] ?? '').trim();
    if (!f || f === 'General' || f === '@') continue;
    counts.set(f, (counts.get(f) ?? 0) + 1);
  }
  return [...counts.entries()]
    .map(([format, count]) => ({ format, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 8);
}

/** Sheet row indexes (0-based) of the formula cells in a column's data rows, capped at 500. */
export function listFormulaRows(formulas: string[][] | undefined, headerRowIndex: number, colIdx: number): number[] {
  const rows: number[] = [];
  if (!formulas) return rows;
  for (let r = headerRowIndex + 1; r < formulas.length && rows.length < 500; r += 1) {
    const f = formulas[r]?.[colIdx];
    if (typeof f === 'string' && f.startsWith('=')) rows.push(r);
  }
  return rows;
}

function sheetToSnapshot(sheet: SheetContext): SheetSnapshot {
  const headerRowIndex = sheet.headerRowIndex ?? 0;
  const headers =
    sheet.headers.length > 0
      ? sheet.headers
      : (sheet.values[headerRowIndex] ?? []).map((v) => String(v ?? ''));
  const sampleData = buildHeadTailSample(sheet.values, headerRowIndex);
  const truncated = sheet.rowCount > sampleData.length + 1;

  const columnMeta: ColumnMeta[] = Array.from(
    { length: Math.max(sheet.columnCount, 1) },
    (_, colIdx) => {
      const header = headers[colIdx] ?? '';
      const colValues = sheet.values.slice(headerRowIndex + 1).map((row) => row[colIdx] ?? null);
      const nonEmpty = colValues.filter((v) => !isBlankCell(v));
      // Prefer the real per-column read (TASKS.md #64, from the first data row) —
      // fall back to the header row's numberFormat only for a minimal/legacy
      // context that never populated columnFormats (e.g. buildMinimalWorkbookContext).
      const columnFormat = sheet.columnFormats?.[colIdx];
      const fmt = columnFormat?.numberFormat ?? String(sheet.numberFormats[headerRowIndex]?.[colIdx] ?? '');
      const hasFontOrFill =
        columnFormat?.bold !== undefined ||
        columnFormat?.italic !== undefined ||
        columnFormat?.fontColor !== undefined ||
        columnFormat?.fillColor !== undefined;

      const formatCounts = countColumnFormats(sheet.numberFormats, headerRowIndex, colIdx);
      const formulaRows = listFormulaRows(sheet.formulas, headerRowIndex, colIdx);

      return {
        index: colIdx,
        header,
        ...(formatCounts.length ? { formatCounts } : {}),
        ...(formulaRows.length ? { formulaRows } : {}),
        sampleValues: nonEmpty.slice(0, 5) as (string | number | null)[],
        detectedType: detectColumnType(nonEmpty, fmt),
        numberFormat:
          detectDateFormat(fmt) ??
          (fmt && fmt !== 'General' && fmt !== '@' ? fmt : undefined),
        ...(hasFontOrFill
          ? {
              format: {
                bold: columnFormat?.bold,
                italic: columnFormat?.italic,
                fontColor: columnFormat?.fontColor,
                fillColor: columnFormat?.fillColor,
              },
            }
          : {}),
      };
    },
  );

  return {
    sheetName: sheet.name,
    usedRange: sheet.usedRange,
    rowCount: sheet.rowCount,
    colCount: Math.max(sheet.columnCount, 1),
    headers,
    headerRowIndex,
    sampleData,
    columnMeta,
    structure: sheet.structure,
    formulaSummary: sheet.formulaSummary,
    isHidden: sheet.isHidden,
    ...(truncated
      ? {
          compressionMeta: {
            originalRowCount: sheet.rowCount,
            compressedRowCount: sampleData.length + (headers.length > 0 ? 1 : 0),
            truncated: true,
            onDemandFetchEnabled: true,
          },
        }
      : {}),
  };
}

export function deepToApiWorkbookContext(deep: DeepWorkbookContext): WorkbookContext {
  return {
    activeSheet: deep.activeSheetName,
    selectedRange: deep.selectedRange,
    sheets: deep.sheets.map(sheetToSnapshot),
    namedRanges: deep.namedRanges,
    tables: deep.tables,
    conditionalFormats: deep.conditionalFormats,
    prompt_context: deep.prompt_context,
  };
}
