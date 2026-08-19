import { describe, expect, it, vi } from 'vitest';
import { handleBatchSet } from './cell.handler';
import type { BatchSetAction } from '@/action.types';

/**
 * Regression: a BATCH_SET action with no operations array reached this
 * handler undiscarded and crashed the entire apply with
 * "action.operations is not iterable" — losing every other already-verified
 * action in the same batch, not just this one. The backend now rejects this
 * shape before it ever reaches the frontend (normalize-executor-output.util.ts),
 * but this handler must not crash on it regardless — defense in depth.
 */
describe('handleBatchSet — missing/malformed operations array', () => {
  function makeCtx() {
    const getItem = vi.fn();
    return {
      ctx: {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext,
      getItem,
    };
  }

  it('resolves without resolving a worksheet when operations is undefined', async () => {
    const { ctx, getItem } = makeCtx();
    const action = { type: 'BATCH_SET', sheetName: 'Main' } as unknown as BatchSetAction;

    await expect(handleBatchSet(action, ctx)).resolves.toBeUndefined();
    expect(getItem).not.toHaveBeenCalled();
  });

  it('resolves without resolving a worksheet when operations is an empty array', async () => {
    const { ctx, getItem } = makeCtx();
    const action = {
      type: 'BATCH_SET',
      sheetName: 'Main',
      operations: [],
    } as unknown as BatchSetAction;

    await expect(handleBatchSet(action, ctx)).resolves.toBeUndefined();
    expect(getItem).not.toHaveBeenCalled();
  });

  it('still processes a well-formed BATCH_SET', async () => {
    const range = { formulas: [['']], numberFormat: [['General']], format: { fill: {} }, load: vi.fn() };
    const { ctx, getItem } = makeCtx();
    getItem.mockReturnValue({ getRange: vi.fn(() => range) });

    const action = {
      type: 'BATCH_SET',
      sheetName: 'Main',
      operations: [{ address: 'B2', formula: '=SUM(January!G:G)' }],
    } as unknown as BatchSetAction;

    await handleBatchSet(action, ctx);
    expect(getItem).toHaveBeenCalledWith('Main');
    expect(range.formulas).toEqual([['=SUM(January!G:G)']]);
  });
});
