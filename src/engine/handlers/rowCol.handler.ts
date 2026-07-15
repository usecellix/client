import {
  AddRowAction,
  DeleteRowAction,
  InsertColumnAction,
  DeleteColumnAction,
} from '@/action.types';
import { resolveWorksheet } from '../sheetResolve';

/* global Excel */

export async function handleAddRow(action: AddRowAction, ctx: Excel.RequestContext): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const insertRow = action.afterRow + 1;
  const insertRange = sheet.getRange(`${insertRow}:${insertRow}`);
  insertRange.insert('Down');

  const colCount = Math.max(action.values.length, 1);
  const newRowRange = sheet.getRangeByIndexes(insertRow - 1, 0, 1, colCount);
  newRowRange.values = [action.values];

  if (action.copyFormatFromRow) {
    const sourceRow = sheet.getRange(`${action.copyFormatFromRow}:${action.copyFormatFromRow}`);
    const targetRow = sheet.getRange(`${insertRow}:${insertRow}`);
    targetRow.copyFrom(sourceRow, Excel.RangeCopyType.formats, false, false);
  }

  await ctx.sync();
}

export async function handleDeleteRow(
  action: DeleteRowAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const sorted = [...action.rows].sort((a, b) => b - a);
  for (const row of sorted) {
    sheet.getRange(`${row}:${row}`).delete('Up');
  }
  await ctx.sync();
}

export async function handleInsertColumn(
  action: InsertColumnAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const col = action.beforeColumn;
  for (let i = 0; i < action.count; i += 1) {
    sheet.getRange(`${col}:${col}`).insert('Right');
  }
  await ctx.sync();
}

export async function handleDeleteColumn(
  action: DeleteColumnAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const sorted = [...action.columns].sort().reverse();
  for (const col of sorted) {
    sheet.getRange(`${col}:${col}`).delete('Left');
  }
  await ctx.sync();
}
