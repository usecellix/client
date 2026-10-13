import { describe, expect, it } from 'vitest';
import type { SheetAction } from '@/types/sheet-actions';
import {
  buildMissingSheetRetry,
  describeMissingSheetFailure,
  parseItemNotFoundFailures,
} from '@/utils/missingSheetRetry';

const NOT_FOUND = "The requested resource doesn't exist.";

// The live step 4 (TASKS.md #346): month-sheet formulas and a Main dashboard,
// with no step anywhere having created Main.
const step4: SheetAction[] = [
  { type: 'SET_FORMULA', sheetName: 'January', range: 'F2:L2', formula: '=E2-D2' },
  { type: 'DATA_VALIDATION', sheetName: 'February', range: 'I2' },
  { type: 'BATCH_SET', sheetName: 'Main', operations: [] },
  { type: 'FORMAT_RANGE', sheetName: 'Main', range: 'A1:E1' },
  { type: 'SET_COLUMN_WIDTH', sheetName: 'January', col: 0, width: 80 },
];
const months = ['January', 'February', 'Sheet1'];

describe('parseItemNotFoundFailures', () => {
  it('returns the failed action types when every failure is ItemNotFound', () => {
    expect(
      parseItemNotFoundFailures(`BATCH_SET: ${NOT_FOUND}; FORMAT_RANGE: ${NOT_FOUND}`),
    ).toEqual(['BATCH_SET', 'FORMAT_RANGE']);
  });

  it('refuses when any other failure is mixed in', () => {
    expect(
      parseItemNotFoundFailures(`BATCH_SET: ${NOT_FOUND}; SET_FORMULA: Invalid argument`),
    ).toBeNull();
  });

  it('refuses errors without an action-type prefix (the preview path, or a wrapper)', () => {
    expect(parseItemNotFoundFailures(NOT_FOUND)).toBeNull();
    expect(parseItemNotFoundFailures(`Spreadsheet update failed: ${NOT_FOUND}`)).toBeNull();
    expect(parseItemNotFoundFailures('')).toBeNull();
  });
});

describe('buildMissingSheetRetry', () => {
  it('creates the missing sheet and re-applies only the actions that touch it', () => {
    const plan = buildMissingSheetRetry(
      step4,
      months,
      `BATCH_SET: ${NOT_FOUND}; FORMAT_RANGE: ${NOT_FOUND}`,
    );
    expect(plan?.missingSheets).toEqual(['Main']);
    expect(plan?.actions.map((a) => `${a.type}:${a.sheetName}`)).toEqual([
      'ADD_SHEET:Main',
      'BATCH_SET:Main',
      'FORMAT_RANGE:Main',
    ]);
    expect(plan?.actions[0].name).toBe('Main');
  });

  it('never replays writes to sheets that exist — those already applied', () => {
    const plan = buildMissingSheetRetry(step4, months, `BATCH_SET: ${NOT_FOUND}`);
    expect(plan?.actions.some((a) => a.sheetName === 'January' || a.sheetName === 'February')).toBe(
      false,
    );
  });

  it("reuses the step's own create instead of adding a second one", () => {
    const withCreate: SheetAction[] = [{ type: 'ADD_SHEET', name: 'Main' }, ...step4];
    const plan = buildMissingSheetRetry(withCreate, months, `BATCH_SET: ${NOT_FOUND}`);
    expect(plan?.actions.filter((a) => a.type === 'ADD_SHEET')).toHaveLength(1);
    expect(plan?.actions[0]).toBe(withCreate[0]);
  });

  it('matches sheet names case-insensitively', () => {
    const plan = buildMissingSheetRetry(
      step4,
      ['JANUARY', 'february'],
      `BATCH_SET: ${NOT_FOUND}`,
    );
    expect(plan?.missingSheets).toEqual(['Main']);
  });

  it('counts cross-sheet sources and destinations as touching a sheet', () => {
    const copy: SheetAction[] = [
      { type: 'COPY_FILTERED_RANGE', sourceSheet: 'January', destSheet: 'Archive' },
    ];
    const plan = buildMissingSheetRetry(copy, months, `COPY_FILTERED_RANGE: ${NOT_FOUND}`);
    expect(plan?.missingSheets).toEqual(['Archive']);
    expect(plan?.actions.map((a) => a.type)).toEqual(['ADD_SHEET', 'COPY_FILTERED_RANGE']);
  });

  it('declines when every sheet exists — the ItemNotFound was something else', () => {
    expect(
      buildMissingSheetRetry(step4, [...months, 'Main'], `BATCH_SET: ${NOT_FOUND}`),
    ).toBeNull();
  });

  it('declines when a failed action type is not one the retry would re-apply', () => {
    // CREATE_CHART failed on an existing sheet (a missing source range, say);
    // creating Main would not fix it.
    const withChart: SheetAction[] = [
      ...step4,
      { type: 'CREATE_CHART', sheetName: 'January', sourceRange: 'A1:B2' },
    ];
    expect(
      buildMissingSheetRetry(
        withChart,
        months,
        `BATCH_SET: ${NOT_FOUND}; CREATE_CHART: ${NOT_FOUND}`,
      ),
    ).toBeNull();
  });

  it('declines when a non-ItemNotFound failure is mixed in', () => {
    expect(
      buildMissingSheetRetry(step4, months, `BATCH_SET: ${NOT_FOUND}; SET_FORMULA: Invalid argument`),
    ).toBeNull();
  });
});

describe('describeMissingSheetFailure', () => {
  it('names the sheet and does not tell the user to re-run the request', () => {
    const message = describeMissingSheetFailure(['Main']);
    expect(message).toContain("'Main'");
    expect(message).not.toMatch(/re-?run/i);
  });

  it('lists several sheets', () => {
    expect(describeMissingSheetFailure(['Main', 'Totals', 'Lists'])).toContain(
      "'Main', 'Totals' and 'Lists'",
    );
  });
});
