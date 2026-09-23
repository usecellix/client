import { describe, expect, it } from 'vitest';
import { parseRangeAddress } from './addressUtils';
import { boundsToAddress, resolveActionSelectBounds } from './selectRanges';
import { SheetAction } from '@/types/sheet-actions';

describe('selectRanges', () => {
  it('parses used-range addresses with sheet prefixes for column placement', () => {
    expect(parseRangeAddress("'Purchase Register'!A1:L52")).toEqual({
      row: 0,
      col: 0,
      rowCount: 52,
      colCount: 12,
    });
    // Next empty column after L is index 12 → M
    const bounds = parseRangeAddress('A1:L52')!;
    expect(bounds.col + bounds.colCount).toBe(12);
  });

  it('resolves INSERT_COLUMN select bounds from stamped metadata', () => {
    const action = {
      type: 'INSERT_COLUMN' as const,
      sheetName: 'Purchase Register',
      columnName: 'Net of Tax',
      position: 'afterLastColumn' as const,
      __appliedBounds: {
        sheetName: 'Purchase Register',
        row: 0,
        col: 12,
        rowCount: 52,
        colCount: 1,
      },
    };
    const bounds = resolveActionSelectBounds(action as never, 'Purchase Register');
    expect(bounds).toEqual({
      sheetName: 'Purchase Register',
      row: 0,
      col: 12,
      rowCount: 52,
      colCount: 1,
    });
    expect(boundsToAddress(bounds!)).toBe('M1:M52');
  });

  it('resolves SET_FORMULA cell for selection', () => {
    const action: SheetAction = {
      type: 'SET_FORMULA',
      sheetName: 'Purchase Register',
      row: 1,
      col: 12,
      formula: '=J2-I2',
    };
    expect(resolveActionSelectBounds(action, 'Purchase Register')).toEqual({
      sheetName: 'Purchase Register',
      row: 1,
      col: 12,
      rowCount: 1,
      colCount: 1,
    });
  });

  /**
   * TASKS.md #252 — a live "highlight rows below 1 lakh" (CONDITIONAL_FORMAT
   * over A2:J61) left the ENTIRE range mouse-selected after apply, which
   * looked exactly like the whole workbook had been accidentally selected —
   * because `range` here is the rule's whole candidate area, not the
   * (usually much smaller) set of rows that actually matched and changed.
   * FORMAT_MATCHING_ROWS/SET_MATCHING_ROWS share the identical shape.
   */
  it('does not select the whole rule range for CONDITIONAL_FORMAT / FORMAT_MATCHING_ROWS / SET_MATCHING_ROWS (#252)', () => {
    const base = { sheetName: 'Purchase Register', range: 'A2:J61' };
    expect(
      resolveActionSelectBounds(
        { type: 'CONDITIONAL_FORMAT', ...base, rule: { kind: 'formula', formula: '=$E2<100000', format: {} } } as never,
        'Purchase Register',
      ),
    ).toBeNull();
    expect(
      resolveActionSelectBounds(
        { type: 'FORMAT_MATCHING_ROWS', ...base, hasHeaders: true, filter: { column: 'X', operator: 'equals', value: 'Y' }, format: {} } as never,
        'Purchase Register',
      ),
    ).toBeNull();
    expect(
      resolveActionSelectBounds(
        { type: 'SET_MATCHING_ROWS', ...base, hasHeaders: true, targetColumn: 'Status', value: 'X' } as never,
        'Purchase Register',
      ),
    ).toBeNull();
  });
});
