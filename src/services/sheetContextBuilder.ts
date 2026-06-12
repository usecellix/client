import { ColumnMeta, SheetSnapshot, WorkbookContext } from '@/types/cellix.types';
import { CompareResult } from '@/types/cellix.types';
import { detectDateFormat } from '@/services/formatGuard';
import { getCompareEndpoint } from '@/lib/apiConfig';

/* global Excel */

function isBlankCell(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function colIndexToLetter(col: number): string {
  let index = col + 1;
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

function formatUsedRange(rowCount: number, colCount: number): string {
  if (rowCount <= 0 || colCount <= 0) return 'A1';
  const end = `${colIndexToLetter(colCount - 1)}${rowCount}`;
  return rowCount === 1 && colCount === 1 ? 'A1' : `A1:${end}`;
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

function buildSnapshotFromValues(
  sheetName: string,
  values: unknown[][],
  numberFormats: unknown[][],
): SheetSnapshot {
  const rowCount = values.length;
  const colCount = values.reduce((max, row) => Math.max(max, row.length), 0);
  const headers = (values[0] ?? []).map((v) => String(v ?? ''));
  const sampleData = values.slice(1, 11).map((row) =>
    row.map((v) => (v === '' || v == null ? null : (v as string | number))),
  );

  const columnMeta: ColumnMeta[] = Array.from({ length: Math.max(colCount, 1) }, (_, colIdx) => {
    const header = headers[colIdx] ?? '';
    const colValues = values.slice(1).map((row) => row[colIdx] ?? null);
    const nonEmpty = colValues.filter((v) => !isBlankCell(v));
    const fmt = String(numberFormats[0]?.[colIdx] ?? '');

    return {
      index: colIdx,
      header,
      sampleValues: nonEmpty.slice(0, 5) as (string | number | null)[],
      detectedType: detectColumnType(nonEmpty, fmt),
      numberFormat:
        detectDateFormat(fmt) ??
        (fmt && fmt !== 'General' && fmt !== '@' ? fmt : undefined),
    };
  });

  return {
    sheetName,
    usedRange: formatUsedRange(rowCount, colCount),
    rowCount,
    colCount: Math.max(colCount, 1),
    headers,
    sampleData,
    columnMeta,
  };
}

async function readSheetValues(
  worksheet: Excel.Worksheet,
  context: Excel.RequestContext,
): Promise<{ values: unknown[][]; numberFormats: unknown[][] }> {
  const usedRange = worksheet.getUsedRange();
  if (!usedRange) {
    return { values: [], numberFormats: [[]] };
  }

  usedRange.load(['values', 'numberFormat', 'rowCount', 'columnCount']);
  await context.sync();

  return {
    values: usedRange.values ?? [],
    numberFormats: usedRange.numberFormat ?? [[]],
  };
}

export async function buildWorkbookContext(
  selectedSheetNames: string[],
): Promise<{ context: WorkbookContext; activeSheetData: unknown[][] }> {
  const names = selectedSheetNames.length ? selectedSheetNames : [];
  let activeSheetName = '';
  let activeSheetData: unknown[][] = [];
  const snapshots: SheetSnapshot[] = [];

  await Excel.run(async (ctx) => {
    const activeWs = ctx.workbook.worksheets.getActiveWorksheet();
    activeWs.load('name');
    await ctx.sync();
    activeSheetName = activeWs.name;

    const targetNames = names.length ? names : [activeSheetName];

    for (const sheetName of targetNames) {
      const ws = ctx.workbook.worksheets.getItem(sheetName);
      const { values, numberFormats } = await readSheetValues(ws, ctx);
      snapshots.push(buildSnapshotFromValues(sheetName, values, numberFormats));

      if (sheetName === activeSheetName) {
        activeSheetData = values;
      }
    }
  });

  return {
    context: {
      sheets: snapshots,
      activeSheet: activeSheetName,
    },
    activeSheetData,
  };
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

export function compareSheetValuesLocally(
  sheetA: string,
  sheetB: string,
  valuesA: unknown[][],
  valuesB: unknown[][],
): CompareResult {
  const maxRows = Math.max(valuesA.length, valuesB.length);
  const maxCols = Math.max(
    valuesA.reduce((m, r) => Math.max(m, r.length), 0),
    valuesB.reduce((m, r) => Math.max(m, r.length), 0),
  );

  const modifiedCells: CompareResult['modifiedCells'] = [];
  const addedInB: string[] = [];
  const removedInB: string[] = [];

  for (let row = 0; row < maxRows; row += 1) {
    const rowA = valuesA[row];
    const rowB = valuesB[row];
    const hasA = rowA?.some((cell) => !isBlankCell(cell));
    const hasB = rowB?.some((cell) => !isBlankCell(cell));

    if (!hasA && hasB) {
      addedInB.push(`Row ${row + 1}`);
      continue;
    }
    if (hasA && !hasB) {
      removedInB.push(`Row ${row + 1}`);
      continue;
    }

    for (let col = 0; col < maxCols; col += 1) {
      const a = normalizeCell(rowA?.[col]);
      const b = normalizeCell(rowB?.[col]);
      if (a !== b) {
        modifiedCells.push({
          address: `${colIndexToLetter(col)}${row + 1}`,
          valueA: a || '(empty)',
          valueB: b || '(empty)',
        });
      }
    }
  }

  return {
    sheetA,
    sheetB,
    summary: `${modifiedCells.length} cell difference(s), ${addedInB.length} row(s) added in B, ${removedInB.length} row(s) removed in B`,
    differences: [],
    addedInB,
    removedInB,
    modifiedCells,
  };
}

async function readFullSheetValues(sheetName: string): Promise<unknown[][]> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    const { values } = await readSheetValues(ws, ctx);
    return values;
  });
}

export async function compareSheets(
  sheetA: string,
  sheetB: string,
): Promise<CompareResult> {
  const context = await buildWorkbookContext([sheetA, sheetB]);

  try {
    const endpoint = getCompareEndpoint();
    const headers: Record<string, string> = { 'Content-Type': 'application/json' };
    if (endpoint.includes('.ngrok-free.app')) {
      headers['ngrok-skip-browser-warning'] = 'true';
    }

    const response = await fetch(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify({ sheetA, sheetB, context: context.context }),
    });

    if (response.ok) {
      const body = (await response.json()) as CompareResult | { data?: CompareResult };
      return ('data' in body && body.data ? body.data : body) as CompareResult;
    }
  } catch (error) {
    console.warn('[Cellix] Compare API unavailable, using local diff:', error);
  }

  const [valuesA, valuesB] = await Promise.all([
    readFullSheetValues(sheetA),
    readFullSheetValues(sheetB),
  ]);

  return compareSheetValuesLocally(sheetA, sheetB, valuesA, valuesB);
}
