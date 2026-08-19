import {
  DeepWorkbookContext,
  NamedRangeInfo,
  SheetContext,
  TableInfo,
} from './context.types';
import { detectSheetStructure, detectHeaders } from './sheetAnalyzer';
import { summarizeFormulas } from './formulaSummarizer';
import { buildPromptContext } from './contextCompressor';

/* global Excel */

function asStringMatrix(matrix: unknown[][] | undefined): string[][] {
  if (!matrix) return [];
  return matrix.map((row) =>
    (row ?? []).map((cell) => (typeof cell === 'string' ? cell : String(cell ?? ''))),
  );
}

function asValueMatrix(matrix: unknown[][] | undefined): unknown[][] {
  if (!matrix) return [];
  return matrix.map((row) => [...(row ?? [])]);
}

/** Minimal active-sheet read when full workbook extraction fails. */
export async function readActiveSheetMinimal(): Promise<{
  name: string;
  values: unknown[][];
  usedRange: string;
  rowCount: number;
  columnCount: number;
}> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');
    await ctx.sync();

    const usedRange = sheet.getUsedRange();
    if (!usedRange) {
      return {
        name: sheet.name,
        values: [],
        usedRange: '',
        rowCount: 0,
        columnCount: 0,
      };
    }

    try {
      usedRange.load(['values', 'address', 'rowCount', 'columnCount']);
      await ctx.sync();
    } catch {
      return {
        name: sheet.name,
        values: [],
        usedRange: '',
        rowCount: 0,
        columnCount: 0,
      };
    }

    return {
      name: sheet.name,
      values: asValueMatrix(usedRange.values),
      usedRange: usedRange.address ?? '',
      rowCount: usedRange.rowCount,
      columnCount: usedRange.columnCount,
    };
  });
}

export async function buildWorkbookContext(
  filterSheetNames?: string[],
): Promise<DeepWorkbookContext> {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;

    const sheets = wb.worksheets;
    sheets.load(['items/name', 'items/visibility']);

    const activeSheet = wb.worksheets.getActiveWorksheet();
    activeSheet.load('name');

    const selection = wb.getSelectedRange();
    selection.load('address');

    await ctx.sync();

    const activeSheetName = activeSheet.name;
    const targetNames = filterSheetNames?.length
      ? new Set([activeSheetName, ...filterSheetNames])
      : null;

    const namedRanges = await loadNamedRangesSafe(ctx, wb);
    const tableInfos = await loadTablesSafe(ctx, wb);

    const sheetContexts: SheetContext[] = [];
    for (const sheet of sheets.items) {
      if (targetNames && !targetNames.has(sheet.name)) continue;
      try {
        sheetContexts.push(await extractSheetContext(ctx, sheet));
      } catch (err) {
        console.warn(`[Cellix] Failed to read sheet "${sheet.name}":`, err);
        sheetContexts.push(emptySheetContext(sheet.name));
      }
    }

    const prompt_context = buildPromptContext(
      activeSheetName,
      sheetContexts,
      namedRanges,
      tableInfos,
      selection.address,
    );

    return {
      activeSheetName,
      selectedRange: selection.address,
      sheets: sheetContexts,
      namedRanges,
      tables: tableInfos,
      prompt_context,
    };
  });
}

async function loadNamedRangesSafe(
  ctx: Excel.RequestContext,
  wb: Excel.Workbook,
): Promise<NamedRangeInfo[]> {
  try {
    const namedItems = wb.names;
    namedItems.load(['items/name', 'items/formula', 'items/type']);
    await ctx.sync();
    return namedItems.items.map((n) => ({
      name: n.name,
      formula: n.formula,
      type: n.type,
    }));
  } catch (err) {
    console.warn('[Cellix] Named ranges read failed:', err);
    return [];
  }
}

async function loadTablesSafe(
  ctx: Excel.RequestContext,
  wb: Excel.Workbook,
): Promise<TableInfo[]> {
  try {
    const tables = wb.tables;
    tables.load(['items/name', 'items/showHeaders', 'items/worksheet/name']);
    await ctx.sync();

    const tableInfos: TableInfo[] = [];
    for (const table of tables.items) {
      try {
        tableInfos.push(await extractTableInfo(ctx, table));
      } catch (err) {
        console.warn(`[Cellix] Table "${table.name}" read failed:`, err);
      }
    }
    return tableInfos;
  } catch (err) {
    console.warn('[Cellix] Tables read failed:', err);
    return [];
  }
}

async function extractSheetContext(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
): Promise<SheetContext> {
  const isHidden = sheet.visibility !== Excel.SheetVisibility.visible;

  if (isHidden) {
    return {
      name: sheet.name,
      usedRange: '',
      rowCount: 0,
      columnCount: 0,
      values: [],
      formulas: [],
      numberFormats: [],
      structure: 'unknown',
      headers: [],
      headerRowIndex: 0,
      formulaSummary: '',
      isHidden: true,
    };
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    return emptySheetContext(sheet.name);
  }

  try {
    usedRange.load([
      'values',
      'formulas',
      'address',
      'rowCount',
      'columnCount',
      'rowIndex',
      'columnIndex',
    ]);
    await ctx.sync();
  } catch {
    return emptySheetContext(sheet.name);
  }

  const values = asValueMatrix(usedRange.values);
  const formulas = asStringMatrix(usedRange.formulas as unknown[][]);
  const { headers, headerRowIndex } = detectHeaders(values);

  let numberFormats: string[][] = [];
  try {
    // Fetch formats for the DETECTED header row, not always the sheet's first
    // row — sheets with a title row above the table have their real header
    // formats several rows down.
    const headerRowRange = sheet.getRangeByIndexes(
      (usedRange.rowIndex ?? 0) + headerRowIndex,
      usedRange.columnIndex ?? 0,
      1,
      usedRange.columnCount,
    );
    headerRowRange.load('numberFormat');
    await ctx.sync();
    const headerFormats = asStringMatrix(headerRowRange.numberFormat as unknown[][])[0] ?? [];
    numberFormats = values.map((row, rowIdx) =>
      rowIdx === headerRowIndex ? headerFormats : Array(row.length).fill(''),
    );
  } catch {
    numberFormats = values.map((row) => Array(row.length).fill(''));
  }

  const structure = detectSheetStructure(values, formulas);
  const { humanReadable: formulaSummary } = summarizeFormulas(formulas);

  return {
    name: sheet.name,
    usedRange: usedRange.address ?? '',
    rowCount: usedRange.rowCount,
    columnCount: usedRange.columnCount,
    values,
    formulas,
    numberFormats,
    structure,
    headers,
    headerRowIndex,
    formulaSummary,
    isHidden: false,
  };
}

function emptySheetContext(name: string): SheetContext {
  return {
    name,
    usedRange: '',
    rowCount: 0,
    columnCount: 0,
    values: [],
    formulas: [],
    numberFormats: [],
    structure: 'unknown',
    headers: [],
    headerRowIndex: 0,
    formulaSummary: '',
    isHidden: false,
  };
}

async function extractTableInfo(
  ctx: Excel.RequestContext,
  table: Excel.Table,
): Promise<TableInfo> {
  const headerRange = table.getHeaderRowRange();
  headerRange.load('values');
  table.worksheet.load('name');
  const range = table.getRange();
  range.load('address');
  await ctx.sync();

  return {
    name: table.name,
    sheetName: table.worksheet.name,
    range: range.address ?? '',
    hasHeaders: table.showHeaders,
    columnNames: (headerRange.values?.[0] ?? []).map(String),
  };
}
