import { describe, expect, it, vi } from 'vitest';
import { handleBatchSet, handleSetRangeValues } from './cell.handler';
import type { BatchSetAction, SetRangeValuesAction } from '@/action.types';

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

/**
 * TASKS.md #100 — SET_RANGE_VALUES is the fast bulk revert path (e.g. for
 * undoing a SORT_RANGE): it reads the range's current live values in one
 * call, overlays a sparse set of corrections at their real positions, and
 * writes the merged block back in one call — instead of one Office.js round
 * trip per changed cell. This pins the read-overlay-write math: corrections
 * must land at the right relative row/col, and cells with no correction must
 * be left exactly as Excel already had them.
 */
describe('handleSetRangeValues — bulk read-overlay-write', () => {
  function makeMockRange(values: unknown[][], rowIndex: number, columnIndex: number) {
    return {
      values,
      rowIndex,
      columnIndex,
      numberFormat: values.map((row) => row.map(() => 'General')),
      rowCount: values.length,
      columnCount: values[0]?.length ?? 0,
      load: vi.fn(),
    } as unknown as Excel.Range;
  }

  function makeCtx(range: Excel.Range) {
    const getRange = vi.fn(() => range);
    return {
      workbook: {
        worksheets: {
          getItem: vi.fn(() => ({ getRange })),
          getActiveWorksheet: vi.fn(() => ({ getRange })),
        },
      },
      sync: vi.fn(async () => undefined),
    } as unknown as Excel.RequestContext;
  }

  it('overlays sparse corrections onto the live current values at the correct positions', async () => {
    // Range A2:B4 — current (post-sort) state on the sheet.
    const currentValues = [
      ['B', 10],
      ['C', 20],
      ['A', 30],
    ];
    const range = makeMockRange(currentValues, 1, 0); // rowIndex=1 -> starts at row 2
    const ctx = makeCtx(range);

    const action: SetRangeValuesAction = {
      type: 'SET_RANGE_VALUES',
      sheetName: 'Sheet1',
      range: 'A2:B4',
      operations: [
        { address: 'A2', value: 'A' },
        { address: 'B2', value: 30 },
        { address: 'A4', value: 'C' },
        { address: 'B4', value: 20 },
        // B3/A3 intentionally left uncorrected — must stay as current values.
      ],
    };

    await handleSetRangeValues(action, ctx);

    expect(range.values).toEqual([
      ['A', 30],
      ['C', 20],
      ['C', 20],
    ]);
  });

  it('does nothing when operations is empty', async () => {
    const range = makeMockRange([['x']], 0, 0);
    const ctx = makeCtx(range);
    const action: SetRangeValuesAction = {
      type: 'SET_RANGE_VALUES',
      sheetName: 'Sheet1',
      range: 'A1:A1',
      operations: [],
    };

    await handleSetRangeValues(action, ctx);
    expect(range.values).toEqual([['x']]);
  });
});
