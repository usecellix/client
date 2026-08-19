import {
  SetCellAction,
  SetFormulaAction,
  FillDownAction,
  BatchSetAction,
} from '@/action.types';
import {
  preserveNumberFormatsAroundWrite,
  writeRangeValuesPreservingNumberFormat,
} from '@/services/formatGuard';
import { resolveWorksheet } from '../sheetResolve';
import { applyRichFormat } from './format.handler';

/* global Excel */

export async function handleSetCell(action: SetCellAction, ctx: Excel.RequestContext): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const range = sheet.getRange(action.address);
  await writeRangeValuesPreservingNumberFormat(ctx, sheet, range, [[action.value]], {
    explicitNumberFormat: action.format?.numberFormat,
  });
  if (action.format) {
    // Preserve helper already applied numberFormat; apply the rest of the style.
    applyRichFormat(range, { ...action.format, numberFormat: undefined });
  }
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

  await preserveNumberFormatsAroundWrite(range, ctx, () => {
    range.formulas = formulaArray;
  });

  if (action.format) {
    applyRichFormat(range, { ...action.format, numberFormat: undefined });
    if (action.format.numberFormat) {
      range.numberFormat = [[action.format.numberFormat]];
    }
  }
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
  // Backend now rejects a BATCH_SET with no operations before it reaches here
  // (normalize-executor-output.util.ts) — defense in depth, since an unguarded
  // iteration here previously crashed the whole apply mid-batch.
  if (!Array.isArray(action.operations) || action.operations.length === 0) return;

  const sheet = resolveWorksheet(ctx, action.sheetName);

  for (const op of action.operations) {
    const range = sheet.getRange(op.address);
    if (op.formula !== undefined) {
      await preserveNumberFormatsAroundWrite(range, ctx, () => {
        range.formulas = [[op.formula!]];
      });
    } else if (op.value !== undefined) {
      await writeRangeValuesPreservingNumberFormat(ctx, sheet, range, [[op.value]], {
        explicitNumberFormat: op.format?.numberFormat,
      });
    }
    if (op.format) {
      applyRichFormat(range, { ...op.format, numberFormat: undefined });
      if (op.formula !== undefined && op.format.numberFormat) {
        range.numberFormat = [[op.format.numberFormat]];
      }
    }
  }
  await ctx.sync();
}
