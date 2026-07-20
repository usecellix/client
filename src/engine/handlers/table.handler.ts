import {
  CreateTableAction,
  DefineNamedRangeAction,
  AutoFitColumnsAction,
} from '@/action.types';
import { resolveWorksheet } from '../sheetResolve';

/* global Excel */

export async function handleCreateTable(
  action: CreateTableAction,
  ctx: Excel.RequestContext,
): Promise<void> {
<<<<<<< HEAD
  const tableName = action.tableName.trim();
  if (!tableName) {
    throw new Error('CREATE_TABLE requires a non-empty tableName');
  }

  const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
  const existing = sheet.tables.getItemOrNullObject(tableName);
  existing.load('isNullObject');
  await ctx.sync();
  if (!existing.isNullObject) return;

=======
  const sheet = resolveWorksheet(ctx, action.sheetName);
>>>>>>> dc51072bf62a34d903b3ceaea1dcb4cf10eb1649
  const table = sheet.tables.add(action.range, action.hasHeaders);
  table.name = tableName;
  if (action.style) table.style = action.style;
  await ctx.sync();
}

export async function handleDefineNamedRange(
  action: DefineNamedRangeAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const names = ctx.workbook.names;
  names.load('items/name');
  await ctx.sync();

  const existing = names.items.find((n) => n.name === action.name);
  if (existing) {
    existing.formula = action.formula;
    if (action.comment) existing.comment = action.comment;
  } else {
    names.add(action.name, action.formula, action.comment);
  }
  await ctx.sync();
}

export async function handleAutofitColumns(
  action: AutoFitColumnsAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  if (action.columns && action.columns.length > 0) {
    for (const col of action.columns) {
      sheet.getRange(`${col}:${col}`).format.autofitColumns();
    }
  } else {
    const used = sheet.getUsedRange();
    if (used) {
      used.format.autofitColumns();
    }
  }
  await ctx.sync();
}
