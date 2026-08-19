import { describe, expect, it, vi } from 'vitest';
import {
  annotateDestOverwriteForCreatedSheets,
  guardAgainstOverwrite,
  OVERWRITE_GUARDED_ACTION_TYPES,
  pruneSpuriousAddSheets,
  rangeHasExistingData,
  shouldGuardSheetActionType,
} from './overwriteGuard';
import { convertLegacyToRich } from './legacyConverter';
import { SheetAction } from '@/types/sheet-actions';
import type { RichAction } from '@/action.types';

describe('overwriteGuard', () => {
  it('detects non-empty cells including Payment Status values', () => {
    expect(rangeHasExistingData([['Paid'], ['Pending'], ['Paid']])).toBe(true);
    expect(rangeHasExistingData([[''], [null], [undefined]])).toBe(false);
    expect(rangeHasExistingData([[0], [false]])).toBe(true);
    expect(rangeHasExistingData([])).toBe(false);
  });

  it('guards every value-writing action type from the incident surface', () => {
    const required = [
      'SET_CELL',
      'SET_FORMULA',
      'WRITE_TABLE',
      'BATCH_SET',
      'FILL_DOWN',
      'COPY_FILTERED_RANGE',
      'MOVE_RANGE',
      'AGGREGATE_TABLE',
      'AUTO_FILL',
      'APPEND_ROW',
      'SET_MATCHING_ROWS',
    ];
    for (const type of required) {
      expect(shouldGuardSheetActionType(type)).toBe(true);
      expect(OVERWRITE_GUARDED_ACTION_TYPES.has(type as never)).toBe(true);
    }
  });

  it('does not guard format-only or structural-shift actions', () => {
    expect(shouldGuardSheetActionType('FORMAT_RANGE')).toBe(false);
    expect(shouldGuardSheetActionType('SORT_RANGE')).toBe(false);
    expect(shouldGuardSheetActionType('DELETE_COLUMN')).toBe(false);
    expect(shouldGuardSheetActionType('HIGHLIGHT_CELL')).toBe(false);
    expect(shouldGuardSheetActionType('INSERT_COLUMN')).toBe(false);
  });

  it('blocks the exact repro pattern: SET_FORMULA into occupied column K', () => {
    const badWrites: SheetAction[] = Array.from({ length: 3 }, (_, i) => ({
      type: 'SET_FORMULA' as const,
      sheetName: 'Purchase Register',
      row: i + 1,
      col: 10,
      formula: `=J${i + 2}-I${i + 2}`,
    }));

    const rich = badWrites.map((a) => convertLegacyToRich(a));
    for (const action of rich) {
      expect(action).not.toBeNull();
      if (!action || Array.isArray(action)) continue;
      expect(action.type).toBe('SET_FORMULA');
      if (action.type === 'SET_FORMULA') {
        expect(action.address).toMatch(/^K\d+$/);
        expect(action.explicitOverwriteConfirmed).toBeFalsy();
        expect(shouldGuardSheetActionType(action.type)).toBe(true);
      }
    }

    expect(rangeHasExistingData([['Paid'], ['Pending'], ['Paid']])).toBe(true);
  });

  it('converts semantic INSERT_COLUMN without a guessed col index', () => {
    const action: SheetAction = {
      type: 'INSERT_COLUMN',
      sheetName: 'Purchase Register',
      columnName: 'Net of Tax',
      position: 'afterLastColumn',
      formula: '=J{row}-I{row}',
    };
    const rich = convertLegacyToRich(action);
    expect(rich).toEqual({
      type: 'INSERT_COLUMN',
      sheetName: 'Purchase Register',
      columnName: 'Net of Tax',
      position: 'afterLastColumn',
      afterColumn: undefined,
      formula: '=J{row}-I{row}',
      explicitOverwriteConfirmed: undefined,
    });
  });

  it('preserves explicitOverwriteConfirmed on revert-style SET_CELL', () => {
    const rich = convertLegacyToRich({
      type: 'SET_CELL',
      sheetName: 'Purchase Register',
      row: 1,
      col: 10,
      value: 'Paid',
      explicitOverwriteConfirmed: true,
    });
    expect(rich).toMatchObject({
      type: 'SET_CELL',
      address: 'K2',
      value: 'Paid',
      explicitOverwriteConfirmed: true,
    });
  });

  it('drops phantom Sheet2 and marks COPY dest overwrite for created sheets', () => {
    const actions = [
      { type: 'ADD_SHEET' as const, name: 'Sheet2' },
      { type: 'ADD_SHEET' as const, name: 'paid paid purchases' },
      {
        type: 'COPY_FILTERED_RANGE' as const,
        destSheet: 'paid paid purchases',
        destStartCell: 'A1',
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:L51',
        hasHeaders: true,
        mode: 'copy' as const,
      },
    ];
    const pruned = pruneSpuriousAddSheets(actions);
    expect(pruned.some((a) => a.name === 'Sheet2')).toBe(false);
    expect(pruned).toHaveLength(2);
    const annotated = annotateDestOverwriteForCreatedSheets(pruned);
    expect(annotated.find((a) => a.type === 'COPY_FILTERED_RANGE')).toMatchObject({
      explicitOverwriteConfirmed: true,
    });
  });
});

describe('guardAgainstOverwrite — BATCH_SET with a missing/malformed operations array', () => {
  // Regression: the backend now rejects a BATCH_SET with no operations array
  // before it ever reaches the frontend (normalize-executor-output.util.ts),
  // but this is the defense-in-depth layer — an unguarded `for...of` here
  // previously threw "action.operations is not iterable" and aborted the
  // entire apply, including every other already-verified action in the batch.
  function makeCtx() {
    const getItem = vi.fn();
    return {
      ctx: {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
      } as unknown as Excel.RequestContext,
      getItem,
    };
  }

  it('does not resolve a worksheet or throw when operations is undefined', async () => {
    const { ctx, getItem } = makeCtx();
    const action = { type: 'BATCH_SET', sheetName: 'Main' } as unknown as RichAction;

    await expect(guardAgainstOverwrite(action, ctx)).resolves.toBeUndefined();
    expect(getItem).not.toHaveBeenCalled();
  });

  it('does not resolve a worksheet or throw when operations is an empty array', async () => {
    const { ctx, getItem } = makeCtx();
    const action = {
      type: 'BATCH_SET',
      sheetName: 'Main',
      operations: [],
    } as unknown as RichAction;

    await expect(guardAgainstOverwrite(action, ctx)).resolves.toBeUndefined();
    expect(getItem).not.toHaveBeenCalled();
  });
});
