import { AddSheetAction, DeleteSheetAction, RenameSheetAction, CopySheetAction } from '@/action.types';

/* global Excel */

export async function handleDeleteSheet(
  action: DeleteSheetAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  ctx.workbook.worksheets.getItem(action.sheetName).delete();
  await ctx.sync();
}

export async function handleAddSheet(action: AddSheetAction, ctx: Excel.RequestContext): Promise<void> {
  const sheets = ctx.workbook.worksheets;
  const existing = sheets.getItemOrNullObject(action.name);
  existing.load('isNullObject');
  await ctx.sync();
  if (!existing.isNullObject) {
    existing.activate();
    await ctx.sync();
    return;
  }

  let created: Excel.Worksheet;

  if (action.copyFrom) {
    const source = sheets.getItem(action.copyFrom);
    const copy = source.copy();
    copy.name = action.name;
    if (action.position !== undefined) {
      copy.position = action.position;
    }
    created = copy;
  } else {
    created = sheets.add(action.name);
    if (action.position !== undefined) {
      created.position = action.position;
    }
  }

  created.activate();
  await ctx.sync();
}

export async function handleRenameSheet(
  action: RenameSheetAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = ctx.workbook.worksheets.getItem(action.oldName);
  sheet.name = action.newName;
  await ctx.sync();
}

export async function handleCopySheet(
  action: CopySheetAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const source = ctx.workbook.worksheets.getItem(action.sourceName);
  const copy = source.copy();
  copy.name = action.newName;
  if (action.position !== undefined) {
    copy.position = action.position;
  }
  await ctx.sync();
}
