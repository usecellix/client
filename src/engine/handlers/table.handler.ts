import {
  CreateTableAction,
  DefineNamedRangeAction,
  AutoFitColumnsAction,
} from '@/action.types';

/* global Excel */

export async function handleCreateTable(
  action: CreateTableAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
  const table = sheet.tables.add(action.range, action.hasHeaders);
  table.name = action.tableName;
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
  const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
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
