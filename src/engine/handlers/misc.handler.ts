import {
  AppendRowAction,
  AutoFillAction,
  ClearRangeAction,
  HighlightCellAction,
  InsertRowAction,
  MergeCellsAction,
  WriteTableAction,
} from '@/action.types';
import { applyFormatGuard, coerceRowDataToReferenceFormats } from '@/services/formatGuard';
import { parseCellAddress } from '../addressUtils';
import { resolveWorksheet } from '../sheetResolve';
import { stampAppliedBounds } from '../selectRanges';

/* global Excel */

const PREVIEW_FILL = '#DCFCE7';

async function getAppendRowIndex(
  worksheet: Excel.Worksheet,
  ctx: Excel.RequestContext,
): Promise<number> {
  const used = worksheet.getUsedRange();
  if (!used) return 0;
  used.load(['rowCount']);
  await ctx.sync();
  return used.rowCount;
}

async function getUsedColumnCount(
  worksheet: Excel.Worksheet,
  ctx: Excel.RequestContext,
): Promise<number> {
  const used = worksheet.getUsedRange();
  if (!used) return 1;
  used.load(['columnCount']);
  await ctx.sync();
  return Math.max(used.columnCount, 1);
}

async function getUsedRowCount(
  worksheet: Excel.Worksheet,
  ctx: Excel.RequestContext,
): Promise<number> {
  const used = worksheet.getUsedRange();
  if (!used) return 1;
  used.load(['rowCount']);
  await ctx.sync();
  return used.rowCount;
}

function asExcelCellValue(value: unknown): string | number | boolean {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return value;
  }
  return String(value);
}

export async function handleAppendRow(
  action: AppendRowAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = resolveWorksheet(ctx, action.sheetName);
  const targetRow = await getAppendRowIndex(worksheet, ctx);
  const colCount = Math.max(action.values.length, await getUsedColumnCount(worksheet, ctx));
  const rowRange = worksheet.getRangeByIndexes(targetRow, 0, 1, colCount);
  const coercedData = await coerceRowDataToReferenceFormats(
    ctx,
    worksheet,
    targetRow,
    'above',
    action.values,
    0,
    colCount,
  );
  rowRange.values = [coercedData];
  await applyFormatGuard(
    ctx,
    worksheet,
    { type: 'ADD_ROW', sheetName: action.sheetName },
    targetRow,
    0,
    1,
    colCount,
  );
  stampAppliedBounds(action, {
    sheetName: action.sheetName,
    row: targetRow,
    col: 0,
    rowCount: 1,
    colCount,
  });
  await ctx.sync();
}

export async function handleInsertRow(
  action: InsertRowAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = resolveWorksheet(ctx, action.sheetName);
  const count = action.count ?? 1;
  const colCount = await getUsedColumnCount(worksheet, ctx);
  // InsertShiftDirection is only Down|Right — "Up" was invalid and threw. Inserting
  // at index N always shifts down, which places the new rows ABOVE row N, so
  // "below" is expressed by inserting one row further down instead.
  const row = action.position === 'below' ? action.row + 1 : action.row;
  const range = worksheet.getRangeByIndexes(row, 0, count, colCount);
  range.getEntireRow().insert('Down');
  await applyFormatGuard(
    ctx,
    worksheet,
    { type: 'INSERT_ROW', sheetName: action.sheetName },
    row,
    0,
    count,
    colCount,
  );
  await ctx.sync();
}

export async function handleAutoFill(
  action: AutoFillAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = resolveWorksheet(ctx, action.sheetName);
  const source = worksheet.getRange(action.startAddress);
  const startCell = parseCellAddress(action.startAddress);
  if (!startCell) return;

  if (action.direction === 'down') {
    const endRow = action.endRow ?? (await getUsedRowCount(worksheet, ctx));
    if (endRow <= startCell.row + 1) return;
    const dest = worksheet.getRangeByIndexes(
      startCell.row,
      startCell.col,
      endRow - startCell.row,
      1,
    );
    source.autoFill(dest, Excel.AutoFillType.fillDefault);
  } else {
    const endCol = action.endCol ?? (await getUsedColumnCount(worksheet, ctx)) - 1;
    if (endCol <= startCell.col) return;
    const dest = worksheet.getRangeByIndexes(
      startCell.row,
      startCell.col,
      1,
      endCol - startCell.col + 1,
    );
    source.autoFill(dest, Excel.AutoFillType.fillDefault);
  }
  await ctx.sync();
}

export async function handleWriteTable(
  action: WriteTableAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  // Server often omits sheetName for "create columns on this sheet" → use active sheet.
  const worksheet = resolveWorksheet(ctx, action.sheetName);
  const headers = action.headers;
  const rows = action.rows;
  if (!headers.length) {
    throw new Error('WRITE_TABLE requires headers');
  }

  const tableRows = [headers, ...rows];
  const rowCount = tableRows.length;
  const colCount = Math.max(
    headers.length,
    ...rows.map((row) => (Array.isArray(row) ? row.length : 0)),
    1,
  );

  // Rectangular primitive matrix only — jagged/object cells cause RichApi 0x80070057.
  const matrix = tableRows.map((row) =>
    Array.from({ length: colCount }, (_, index) =>
      asExcelCellValue(Array.isArray(row) ? row[index] : ''),
    ),
  );

  const range = worksheet.getRangeByIndexes(0, 0, rowCount, colCount);
  range.values = matrix;
  await ctx.sync();
}

export async function handleHighlightCell(
  action: HighlightCellAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = resolveWorksheet(ctx, action.sheetName);
  const range = worksheet.getRange(action.address);
  range.format.fill.color = action.color ?? PREVIEW_FILL;
  await ctx.sync();
}

export async function handleMergeCells(
  action: MergeCellsAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = resolveWorksheet(ctx, action.sheetName);
  worksheet.getRange(action.range).merge(false);
  await ctx.sync();
}

export async function handleClearRange(
  action: ClearRangeAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = resolveWorksheet(ctx, action.sheetName);
  const range = worksheet.getRange(action.range);
  const applyTo =
    action.mode === 'contents'
      ? Excel.ClearApplyTo.contents
      : action.mode === 'formats'
        ? Excel.ClearApplyTo.formats
        : Excel.ClearApplyTo.all;
  range.clear(applyTo);
  await ctx.sync();
}
