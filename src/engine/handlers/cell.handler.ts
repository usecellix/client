import {
  SetCellAction,
  SetFormulaAction,
  FillDownAction,
  BatchSetAction,
  SetRangeValuesAction,
} from '@/action.types';
import {
  preserveNumberFormatsAroundWrite,
  writeRangeValuesPreservingNumberFormat,
} from '@/services/formatGuard';
import { resolveWorksheet } from '../sheetResolve';
import { applyRichFormat } from './format.handler';
import { parseCellAddress, parseRangeAddress } from '../addressUtils';

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

/**
 * Revert-only bulk write (TASKS.md #100) — the fast path for reverting an
 * action (e.g. `SORT_RANGE`) that can touch hundreds of cells at once,
 * where a per-cell `SET_CELL` inverse means hundreds of separate Office.js
 * round trips. `action.operations` is deliberately sparse (only the cells
 * that actually need correcting, not every cell in `action.range`) — the
 * backend that builds this action has real captured values for the changed
 * cells but no way to know the unchanged ones, so this reads the range's
 * *current* live values in one call, overlays the corrections in memory,
 * and writes the merged block back in one call. Unlisted cells inside the
 * range are left exactly as Excel already has them. Uses
 * `preserveNumberFormatsAroundWrite` with no remap (values go back to the
 * exact positions they came from, no reordering) so restoring a date
 * column doesn't fall into the same smart-entry reformatting bug class
 * `sort.handler.ts` guards against.
 */
export async function handleSetRangeValues(
  action: SetRangeValuesAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  if (!action.operations.length) return;

  const sheet = resolveWorksheet(ctx, action.sheetName);
  const range = sheet.getRange(action.range);
  range.load(['values', 'rowIndex', 'columnIndex']);
  await ctx.sync();

  const bounds = parseRangeAddress(action.range);
  const baseRow = bounds?.row ?? range.rowIndex;
  const baseCol = bounds?.col ?? range.columnIndex;
  const matrix = (range.values as unknown[][]).map((row) => [...row]);

  for (const op of action.operations) {
    if (op.value === undefined) continue;
    const cell = parseCellAddress(op.address);
    if (!cell) continue;
    const relRow = cell.row - baseRow;
    const relCol = cell.col - baseCol;
    if (relRow < 0 || relCol < 0 || relRow >= matrix.length || relCol >= (matrix[0]?.length ?? 0)) {
      continue;
    }
    matrix[relRow]![relCol] = op.value;
  }

  await preserveNumberFormatsAroundWrite(range, ctx, () => {
    range.values = matrix;
  });
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
