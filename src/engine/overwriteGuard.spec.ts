import { describe, expect, it } from 'vitest';
import {
  OVERWRITE_GUARDED_ACTION_TYPES,
  rangeHasExistingData,
  shouldGuardSheetActionType,
} from './overwriteGuard';
import { convertLegacyToRich } from './legacyConverter';
import { SheetAction } from '@/types/sheet-actions';

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
    // Reconstruct: 12-col sheet, Executor wrongly targets col 10 (K = Payment Status)
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

    // Simulated live values for K2:K4 — guard must treat as occupied
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
});
