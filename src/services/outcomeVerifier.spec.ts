import { describe, expect, it, vi, afterEach } from 'vitest';
import {
  describeOutcome,
  isFormulaError,
  valuesAgree,
  verifyAppliedOutcome,
  verifyAppliedOutcomeSafe,
} from './outcomeVerifier';
import { CellChange } from '@/types/changeSet';

/* global globalThis */

/**
 * TASKS.md #150 — post-apply outcome verification.
 *
 * The cases below are the three real failures `COMPETITIVE_STUDY_SHORTCUT.md`
 * recorded, each of which produced well-formed actions and a wrong workbook, so
 * none was catchable by any amount of pre-write checking:
 *   - `ADD_SHEET "Main"` silently became `Main 2`
 *   - Accept applied nothing at all
 *   - a written formula landed as `#REF!`
 */

function change(sheet: string, cell: string, after: unknown, formula?: string): CellChange {
  return { sheet, cell, before: '', after, isHardcoded: false, ...(formula ? { formula } : {}) };
}

/** `live[sheet][cell]` — absent sheet key means the sheet does not exist. */
function installExcelMock(live: Record<string, Record<string, unknown>>) {
  const getItemOrNullObject = vi.fn((name: string) => {
    const sheet = live[name];
    return {
      isNullObject: sheet === undefined,
      load: vi.fn(),
      getRange: (cell: string) => ({
        load: vi.fn(),
        values: [[sheet?.[cell]]],
      }),
    };
  });

  (globalThis as Record<string, unknown>).Excel = {
    run: async (cb: (ctx: unknown) => Promise<void>) =>
      cb({
        workbook: { worksheets: { getItemOrNullObject } },
        sync: vi.fn(async () => undefined),
      }),
  };
  return { getItemOrNullObject };
}

describe('valuesAgree — lenient about representation, strict about meaning', () => {
  it('treats numeric and string forms of the same number as equal', () => {
    expect(valuesAgree(0, '0')).toBe(true);
    expect(valuesAgree('1500', 1500)).toBe(true);
    expect(valuesAgree(1.0000000001, 1)).toBe(true);
  });

  it('treats blank forms as equal', () => {
    expect(valuesAgree('', null)).toBe(true);
    expect(valuesAgree(undefined, '')).toBe(true);
  });

  it('still catches a genuinely different value', () => {
    expect(valuesAgree('Unit No', 'Dashboard')).toBe(false);
    expect(valuesAgree(100, 200)).toBe(false);
  });
});

describe('isFormulaError', () => {
  it('recognises the Excel error literals a bad formula produces', () => {
    for (const err of ['#REF!', '#NAME?', '#VALUE!', '#SPILL!', '#DIV/0!', '#N/A']) {
      expect(isFormulaError(err)).toBe(true);
    }
  });

  it('does not flag ordinary content', () => {
    expect(isFormulaError('Unit No')).toBe(false);
    expect(isFormulaError(0)).toBe(false);
    expect(isFormulaError('#hashtag')).toBe(false);
  });
});

describe('verifyAppliedOutcome', () => {
  const originalExcel = (globalThis as Record<string, unknown>).Excel;
  afterEach(() => {
    (globalThis as Record<string, unknown>).Excel = originalExcel;
  });

  it('verifies a clean apply silently', async () => {
    installExcelMock({ January: { A1: 'Unit No', B1: 'Guest' } });
    const result = await verifyAppliedOutcome([
      change('January', 'A1', 'Unit No'),
      change('January', 'B1', 'Guest'),
    ]);

    expect(result.verified).toBe(2);
    expect(result.mismatches).toHaveLength(0);
    expect(describeOutcome(result)).toBeNull();
  });

  it('catches the Main -> Main 2 failure: the ChangeSet wrote to a sheet that is not there', async () => {
    installExcelMock({ 'Main 2': { A1: 'Dashboard' } });
    const result = await verifyAppliedOutcome([change('Main', 'A1', 'Dashboard')]);

    expect(result.mismatches).toHaveLength(1);
    expect(result.mismatches[0].actual).toBe('(sheet missing)');
    expect(describeOutcome(result)).toContain('could not be found afterwards');
    expect(describeOutcome(result)).toContain('Main');
  });

  it('catches "Accept applied nothing" — cells still hold their old values', async () => {
    installExcelMock({ January: { A1: '', B1: '' } });
    const result = await verifyAppliedOutcome([
      change('January', 'A1', 'Unit No'),
      change('January', 'B1', 'Guest'),
    ]);

    expect(result.verified).toBe(0);
    expect(result.mismatches).toHaveLength(2);
    expect(describeOutcome(result)).toContain('do not match what was proposed');
  });

  it('catches a formula that landed as #REF!', async () => {
    installExcelMock({ Main: { A19: '#REF!' } });
    const result = await verifyAppliedOutcome([
      change('Main', 'A19', '', '=LET(rows,VSTACK(January!A2:J500),rows)'),
    ]);

    expect(result.mismatches).toHaveLength(1);
    expect(result.mismatches[0].isFormulaError).toBe(true);
    expect(describeOutcome(result)).toContain('returned an error');
    expect(describeOutcome(result)).toContain('#REF!');
  });

  it('catches #NAME? — the pre-365 dynamic-array failure the VSTACK formula risks', async () => {
    installExcelMock({ Main: { A19: '#NAME?' } });
    const result = await verifyAppliedOutcome([
      change('Main', 'A19', '', '=VSTACK(January!A2:J500)'),
    ]);
    expect(result.mismatches[0].isFormulaError).toBe(true);
    expect(describeOutcome(result)).toContain('#NAME?');
  });

  it('passes a formula cell whose computed value differs from its formula text', async () => {
    // The common case: ChangeSet records the formula, Excel returns the result.
    installExcelMock({ Main: { B5: 0 } });
    const result = await verifyAppliedOutcome([
      change('Main', 'B5', '=SUM(January!G:G)', '=SUM(January!G:G)'),
    ]);
    expect(result.verified).toBe(1);
    expect(result.mismatches).toHaveLength(0);
  });

  it('batches reads per sheet, not per cell', async () => {
    const { getItemOrNullObject } = installExcelMock({
      January: { A1: 'a', B1: 'b', C1: 'c' },
      February: { A1: 'a' },
    });
    await verifyAppliedOutcome([
      change('January', 'A1', 'a'),
      change('January', 'B1', 'b'),
      change('January', 'C1', 'c'),
      change('February', 'A1', 'a'),
    ]);
    // Two sheets touched, two lookups — not four.
    expect(getItemOrNullObject).toHaveBeenCalledTimes(2);
  });

  it('reports nothing for an empty change set', async () => {
    installExcelMock({});
    const result = await verifyAppliedOutcome([]);
    expect(result.verified).toBe(0);
    expect(result.skipped).toBe(false);
    expect(describeOutcome(result)).toBeNull();
  });
});

describe('verifyAppliedOutcomeSafe — never turns a real success into a failure', () => {
  const originalExcel = (globalThis as Record<string, unknown>).Excel;
  afterEach(() => {
    (globalThis as Record<string, unknown>).Excel = originalExcel;
  });

  it('reports skipped rather than throwing when the read-back itself fails', async () => {
    (globalThis as Record<string, unknown>).Excel = {
      run: async () => {
        throw new Error('Excel is in cell-editing mode.');
      },
    };
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);

    const result = await verifyAppliedOutcomeSafe([change('January', 'A1', 'x')]);

    expect(result.skipped).toBe(true);
    expect(result.skipReason).toContain('cell-editing mode');
    // Skipped is "unknown", never "clean" — it must not claim a verdict.
    expect(describeOutcome(result)).toBeNull();
    warn.mockRestore();
  });
});
