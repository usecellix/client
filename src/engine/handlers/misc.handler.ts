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

export async function handleAppendRow(
  action: AppendRowAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = ctx.workbook.worksheets.getItem(action.sheetName);
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
  await applyFormatGuard(ctx, worksheet, { type: 'ADD_ROW', sheetName: action.sheetName }, targetRow, 0, 1, colCount);
  await ctx.sync();
}

export async function handleInsertRow(
  action: InsertRowAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = ctx.workbook.worksheets.getItem(action.sheetName);
  const row = action.row;
  const count = action.count ?? 1;
  const colCount = await getUsedColumnCount(worksheet, ctx);
  const range = worksheet.getRangeByIndexes(row, 0, count, colCount);
  range.getEntireRow().insert(action.position === 'above' ? 'Up' : 'Down');
  await applyFormatGuard(ctx, worksheet, { type: 'INSERT_ROW', sheetName: action.sheetName }, row, 0, count, colCount);
  await ctx.sync();
}

export async function handleAutoFill(
  action: AutoFillAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = ctx.workbook.worksheets.getItem(action.sheetName);
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
  const worksheet = ctx.workbook.worksheets.getItem(action.sheetName);
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

  const range = worksheet.getRangeByIndexes(0, 0, rowCount, colCount);
  range.values = tableRows.map((row) =>
    Array.from({ length: colCount }, (_, index) =>
      Array.isArray(row) ? (row[index] ?? '') : '',
    ),
  );
  await ctx.sync();
}

export async function handleHighlightCell(
  action: HighlightCellAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = ctx.workbook.worksheets.getItem(action.sheetName);
  const range = worksheet.getRange(action.address);
  range.format.fill.color = action.color ?? PREVIEW_FILL;
  await ctx.sync();
}

export async function handleMergeCells(
  action: MergeCellsAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = ctx.workbook.worksheets.getItem(action.sheetName);
  worksheet.getRange(action.range).merge(false);
  await ctx.sync();
}

export async function handleClearRange(
  action: ClearRangeAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const worksheet = ctx.workbook.worksheets.getItem(action.sheetName);
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
