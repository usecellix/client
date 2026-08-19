import { describe, expect, it, vi } from 'vitest';
import { handleSortRange } from './sort.handler';
import type { SortRangeAction } from '@/action.types';

/**
 * Regression: user reported sorting a sheet silently mangled a date column —
 * "11/09/2025" became "11092025". The original handler set `range.values`
 * and `range.numberFormat` together in one batch; Excel's own smart-entry
 * parsing can reinterpret a written value (e.g. re-detect a date and apply a
 * locale default) unless the format is explicitly re-asserted *after* the
 * value write completes. This test checks the one thing a mock actually can:
 * that each row's original format value travels with that row to its new
 * position after sorting, regardless of which Excel API calls carry it there.
 */
describe('handleSortRange — format follows its row after sorting', () => {
  function makeMockRange(values: unknown[][], numberFormat: string[][]) {
    return {
      values,
      numberFormat,
      rowCount: values.length,
      columnCount: values[0]?.length ?? 0,
      rowIndex: 0,
      columnIndex: 0,
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

  it('keeps a row\'s date format paired with its date value after reordering', async () => {
    // Header + 3 data rows: Date column (A) with distinct formats per row to
    // make misalignment detectable, Amount column (B) drives the sort.
    const values = [
      ['Date', 'Amount'],
      [45543, 300], // row for Amount 300 — format 'dd/mm/yyyy'
      [45544, 100], // row for Amount 100 — format 'mm/dd/yyyy'
      [45545, 200], // row for Amount 200 — format 'yyyy-mm-dd'
    ];
    const numberFormat = [
      ['General', 'General'],
      ['dd/mm/yyyy', 'General'],
      ['mm/dd/yyyy', 'General'],
      ['yyyy-mm-dd', 'General'],
    ];
    const range = makeMockRange(values, numberFormat);
    const ctx = makeCtx(range);

    const action: SortRangeAction = {
      type: 'SORT_RANGE',
      sheetName: 'Sheet1',
      range: 'A1:B4',
      key: 1, // Amount column
      ascending: true,
      hasHeaders: true,
    };

    await handleSortRange(action, ctx);

    // Ascending by Amount: 100 (was row2, mm/dd/yyyy), 200 (was row3, yyyy-mm-dd), 300 (was row1, dd/mm/yyyy)
    const finalValues = range.values as unknown[][];
    const finalFormats = range.numberFormat as string[][];

    expect(finalValues[1]).toEqual([45544, 100]);
    expect(finalFormats[1]![0]).toBe('mm/dd/yyyy');

    expect(finalValues[2]).toEqual([45545, 200]);
    expect(finalFormats[2]![0]).toBe('yyyy-mm-dd');

    expect(finalValues[3]).toEqual([45543, 300]);
    expect(finalFormats[3]![0]).toBe('dd/mm/yyyy');

    // Header row's own format must be untouched.
    expect(finalFormats[0]).toEqual(['General', 'General']);
  });

  it('re-asserts numberFormat after writing values, not merely alongside it', async () => {
    // The actual fix under test: preserveNumberFormatsAroundWrite snapshots
    // numberFormat, runs the value write, THEN sets numberFormat — as two
    // distinct property assignments in a specific order, not a single
    // simultaneous batch. Confirm that ordering held.
    const values = [
      ['Date'],
      [45543],
    ];
    const numberFormat = [
      ['General'],
      ['dd/mm/yyyy'],
    ];
    const range = makeMockRange(values, numberFormat);
    const ctx = makeCtx(range);
    const assignmentOrder: string[] = [];

    // Intercept both setters to record ordering.
    let currentValues = range.values;
    let currentFormat = range.numberFormat;
    Object.defineProperty(range, 'values', {
      get: () => currentValues,
      set: (v) => {
        assignmentOrder.push('values');
        currentValues = v;
      },
    });
    Object.defineProperty(range, 'numberFormat', {
      get: () => currentFormat,
      set: (v) => {
        assignmentOrder.push('numberFormat');
        currentFormat = v;
      },
    });

    const action: SortRangeAction = {
      type: 'SORT_RANGE',
      sheetName: 'Sheet1',
      range: 'A1:A2',
      key: 0,
      ascending: true,
      hasHeaders: true,
    };

    await handleSortRange(action, ctx);

    const numberFormatIndex = assignmentOrder.indexOf('numberFormat');
    const valuesIndex = assignmentOrder.indexOf('values');
    expect(valuesIndex).toBeGreaterThanOrEqual(0);
    expect(numberFormatIndex).toBeGreaterThan(valuesIndex);
  });
});
