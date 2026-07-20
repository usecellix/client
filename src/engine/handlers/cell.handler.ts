import {
  SetCellAction,
  SetFormulaAction,
  FillDownAction,
  BatchSetAction,
} from '@/action.types';
import { applyRichFormat } from './format.handler';
import { resolveWorksheet } from './resolveWorksheet';

/* global Excel */

export async function handleSetCell(action: SetCellAction, ctx: Excel.RequestContext): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const range = sheet.getRange(action.address);
  range.values = [[action.value]];
  if (action.format) applyRichFormat(range, action.format);
  await ctx.sync();
}

export async function handleSetFormula(
  action: SetFormulaAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const range = sheet.getRange(action.address);
  range.load(['rowCount', 'columnCount']);
  await ctx.sync();

  const formulaArray = Array.from({ length: range.rowCount }, () =>
    Array.from({ length: range.columnCount }, () => action.formula),
  );
  range.formulas = formulaArray;

  if (action.format) applyRichFormat(range, action.format);
  await ctx.sync();
}

export async function handleFillDown(
  action: FillDownAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const source = sheet.getRange(action.sourceRange);
  const target = sheet.getRange(action.targetRange);
  target.copyFrom(source, Excel.RangeCopyType.all, false, false);
  await ctx.sync();
}

export async function handleBatchSet(
  action: BatchSetAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);

  for (const op of action.operations) {
    const range = sheet.getRange(op.address);
    if (op.formula !== undefined) {
      range.formulas = [[op.formula]];
    } else if (op.value !== undefined) {
      range.values = [[op.value]];
    }
    if (op.format) applyRichFormat(range, op.format);
  }

  await ctx.sync();
}
