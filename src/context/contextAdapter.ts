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

function buildHeadTailSample(values: unknown[][], head = 5, tail = 5): (string | number | null)[][] {
  if (values.length <= 1) return [];
  const dataRows = values.slice(1);
  const selected =
    dataRows.length <= head + tail
      ? dataRows
      : [...dataRows.slice(0, head), ...dataRows.slice(-tail)];

  return selected.map((row) =>
    row.map((v) => (v === '' || v == null ? null : (v as string | number))),
  );
}

function sheetToSnapshot(sheet: SheetContext): SheetSnapshot {
  const headers =
    sheet.headers.length > 0
      ? sheet.headers
      : (sheet.values[0] ?? []).map((v) => String(v ?? ''));
  const sampleData = buildHeadTailSample(sheet.values);
  const truncated = sheet.rowCount > sampleData.length + 1;

  const columnMeta: ColumnMeta[] = Array.from(
    { length: Math.max(sheet.columnCount, 1) },
    (_, colIdx) => {
      const header = headers[colIdx] ?? '';
      const colValues = sheet.values.slice(1).map((row) => row[colIdx] ?? null);
      const nonEmpty = colValues.filter((v) => !isBlankCell(v));
      const fmt = String(sheet.numberFormats[0]?.[colIdx] ?? '');

      return {
        index: colIdx,
        header,
        sampleValues: nonEmpty.slice(0, 5) as (string | number | null)[],
        detectedType: detectColumnType(nonEmpty, fmt),
        numberFormat:
          detectDateFormat(fmt) ??
          (fmt && fmt !== 'General' && fmt !== '@' ? fmt : undefined),
      };
    },
  );

  return {
    sheetName: sheet.name,
    usedRange: sheet.usedRange,
    rowCount: sheet.rowCount,
    colCount: Math.max(sheet.columnCount, 1),
    headers,
    sampleData,
    columnMeta,
    structure: sheet.structure,
    formulaSummary: sheet.formulaSummary,
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
    prompt_context: deep.prompt_context,
  };
}
