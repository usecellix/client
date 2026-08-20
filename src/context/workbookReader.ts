import {
  ConditionalFormatRuleInfo,
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
    const conditionalFormats: ConditionalFormatRuleInfo[] = [];
    for (const sheet of sheets.items) {
      if (targetNames && !targetNames.has(sheet.name)) continue;
      try {
        const { sheetContext, conditionalFormats: sheetCFs } = await extractSheetContext(ctx, sheet);
        sheetContexts.push(sheetContext);
        conditionalFormats.push(...sheetCFs);
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
      conditionalFormats,
    );

    return {
      activeSheetName,
      selectedRange: selection.address,
      sheets: sheetContexts,
      namedRanges,
      tables: tableInfos,
      conditionalFormats,
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

/**
 * Column-level bold/italic/fontColor/fillColor + numberFormat, read from ONE
 * representative row per sheet — the first real data row, not the header row
 * (TASKS.md #64). Reading the header row (as the pre-existing numberFormat-only
 * capture did) would be actively wrong for bold/italic: headers are almost
 * always bold, data rows usually aren't, so broadcasting the header's style
 * to every data cell would make revert *introduce* incorrect formatting
 * rather than restore it. Deliberately one row, not the whole used range —
 * `columnCount` cells is cheap regardless of sheet size, unlike a full-range
 * per-cell read, and `getCellProperties` (ExcelApi 1.9) is the only Office.js
 * API that returns font/fill as a per-cell matrix at all (unlike `numberFormat`,
 * `range.format.font.bold` is a single value for the whole range, not an array).
 */
export async function extractColumnFormats(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  usedRange: Excel.Range,
  headerRowIndex: number,
  rowCount: number,
  columnCount: number,
): Promise<import('./context.types').ColumnFormatSnapshot[]> {
  const empty = (): import('./context.types').ColumnFormatSnapshot[] =>
    Array.from({ length: columnCount }, () => ({}));
  if (columnCount <= 0) return [];

  try {
    const sampleRowIndex = headerRowIndex + 1 < rowCount ? headerRowIndex + 1 : headerRowIndex;
    const sampleRow = sheet.getRangeByIndexes(
      (usedRange.rowIndex ?? 0) + sampleRowIndex,
      usedRange.columnIndex ?? 0,
      1,
      columnCount,
    );
    sampleRow.load('numberFormat');
    const cellProps = sampleRow.getCellProperties({
      format: { font: { bold: true, italic: true, color: true }, fill: { color: true } },
    });
    await ctx.sync();

    const numberFormats = asStringMatrix(sampleRow.numberFormat as unknown[][])[0] ?? [];
    const props = cellProps.value?.[0] ?? [];

    return Array.from({ length: columnCount }, (_, col) => {
      const cellFormat = props[col]?.format;
      const fontColor = cellFormat?.font?.color;
      const fillColor = cellFormat?.fill?.color;
      const numberFormat = numberFormats[col];
      return {
        numberFormat: numberFormat && numberFormat.trim() ? numberFormat : undefined,
        bold: cellFormat?.font?.bold,
        italic: cellFormat?.font?.italic,
        // Office.js reports an empty string for "no explicit override" (automatic
        // color) — treat that as "not captured" rather than a literal empty color,
        // matching applyFormat's `!== undefined` gate so revert doesn't try to set
        // font/fill color to '' and either no-op unpredictably or throw.
        fontColor: fontColor && fontColor.trim() ? fontColor : undefined,
        fillColor: fillColor && fillColor.trim() ? fillColor : undefined,
      };
    });
  } catch (err) {
    console.warn('[Cellix] Column format read failed:', err);
    return empty();
  }
}

async function extractSheetContext(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
): Promise<{ sheetContext: SheetContext; conditionalFormats: ConditionalFormatRuleInfo[] }> {
  const isHidden = sheet.visibility !== Excel.SheetVisibility.visible;

  if (isHidden) {
    return {
      sheetContext: {
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
      },
      conditionalFormats: [],
    };
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    return { sheetContext: emptySheetContext(sheet.name), conditionalFormats: [] };
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
    return { sheetContext: emptySheetContext(sheet.name), conditionalFormats: [] };
  }

  const conditionalFormats = await extractConditionalFormats(ctx, sheet.name, usedRange);

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

  const columnFormats = await extractColumnFormats(
    ctx,
    sheet,
    usedRange,
    headerRowIndex,
    usedRange.rowCount,
    usedRange.columnCount,
  );

  const structure = detectSheetStructure(values, formulas);
  const { humanReadable: formulaSummary } = summarizeFormulas(formulas);

  return {
    sheetContext: {
      name: sheet.name,
      usedRange: usedRange.address ?? '',
      rowCount: usedRange.rowCount,
      columnCount: usedRange.columnCount,
      values,
      formulas,
      numberFormats,
      columnFormats,
      structure,
      headers,
      headerRowIndex,
      formulaSummary,
      isHidden: false,
    },
    conditionalFormats,
  };
}

/**
 * Reads back conditional-format rules already on the sheet's used range via
 * Office.js — not limited to rules Cellix itself applied (TASKS.md #38).
 * Presence/summary only, matching the same shape the shadow-workbook tracks
 * (#39) — enough for a follow-up request to target an existing rule by id
 * rather than stacking a duplicate.
 */
async function extractConditionalFormats(
  ctx: Excel.RequestContext,
  sheetName: string,
  usedRange: Excel.Range,
): Promise<ConditionalFormatRuleInfo[]> {
  try {
    const cfs = usedRange.conditionalFormats;
    cfs.load(['items/id', 'items/type']);
    await ctx.sync();

    const items = cfs.items;
    if (items.length === 0) return [];

    const ranges = items.map((cf) => {
      const range = cf.getRangeOrNullObject();
      range.load(['address', 'isNullObject']);
      return range;
    });
    items.forEach((cf) => {
      if (cf.type === 'CellValue') cf.cellValue.load('rule');
      else if (cf.type === 'Custom') cf.custom.load('rule');
      else if (cf.type === 'TopBottom') cf.topBottom.load('rule');
      else if (cf.type === 'ColorScale') cf.colorScale.load('criteria');
    });
    await ctx.sync();

    const infos: ConditionalFormatRuleInfo[] = [];
    items.forEach((cf, i) => {
      const range = ranges[i];
      if (!range || range.isNullObject) return; // multi-range rule — skip, edge case
      const { ruleKind, summary } = summarizeConditionalFormat(cf);
      infos.push({
        id: cf.id,
        sheetName,
        range: range.address ?? '',
        ruleKind,
        summary,
      });
    });
    return infos;
  } catch (err) {
    console.warn(`[Cellix] Conditional-format read failed for "${sheetName}":`, err);
    return [];
  }
}

function summarizeConditionalFormat(
  cf: Excel.ConditionalFormat,
): { ruleKind: ConditionalFormatRuleInfo['ruleKind']; summary: string } {
  try {
    if (cf.type === 'CellValue') {
      const rule = cf.cellValue.rule;
      const range = rule.formula2 ? `${rule.formula1}..${rule.formula2}` : rule.formula1;
      return { ruleKind: 'cellValue', summary: `${rule.operator} ${range}` };
    }
    if (cf.type === 'Custom') {
      return { ruleKind: 'formula', summary: cf.custom.rule.formula };
    }
    if (cf.type === 'TopBottom') {
      const rule = cf.topBottom.rule;
      return { ruleKind: 'topBottom', summary: `${rule.type} ${rule.rank}` };
    }
    if (cf.type === 'ColorScale') {
      const criteria = cf.colorScale.criteria;
      const stops = [criteria.minimum?.color, criteria.midpoint?.color, criteria.maximum?.color]
        .filter((c): c is string => Boolean(c));
      return { ruleKind: 'colorScale', summary: `scale ${stops.join(' -> ')}` };
    }
  } catch {
    // Fall through to 'other' below — a sub-object failing to load shouldn't
    // drop the whole rule from context, just its detail.
  }
  return { ruleKind: 'other', summary: `${cf.type} rule` };
}

/** Re-exported for unit tests */
export { extractConditionalFormats, summarizeConditionalFormat };

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
