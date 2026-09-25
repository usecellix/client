import { describe, expect, it } from 'vitest';
import {
  cellAlreadyHolds,
  guardAgainstOverwrite,
  isOverwriteGuardError,
  preflightOverwriteGuard,
} from './overwriteGuard';
import type { RichAction } from '@/action.types';

/**
 * TASKS.md #319 — live run_1790326084342_hctud2t, step 6 of 7.
 *
 * The card was BATCH_SET (A1 = "Payments Dashboard" + KPIs), formatting,
 * CREATE_CHART, then a SET_FORMULA at A19. The per-action guard blocked A19,
 * but its own `ctx.sync()` had already committed everything before it. The card
 * stayed pending, and every retry was refused at A1 by the step's own title.
 */

type Cell = { value?: unknown; formula?: string };

/** A minimal Office.js double: named sheets of A1 cells, loads resolved on sync. */
function makeWorkbook(sheets: Record<string, Record<string, Cell>>) {
  const makeRange = (cells: Record<string, Cell>, address: string) => {
    const cell = cells[address.toUpperCase()] ?? {};
    const value = cell.formula ? (cell.value ?? '') : (cell.value ?? '');
    return {
      load: () => undefined,
      values: [[value]],
      formulas: [[cell.formula ?? (cell.value ?? '')]],
    };
  };
  const sheetOf = (name: string) => {
    const cells = sheets[name];
    if (!cells) throw new Error(`ItemNotFound: ${name}`);
    return { getRange: (address: string) => makeRange(cells, address) };
  };
  const ctx = {
    workbook: {
      worksheets: {
        getItem: (name: string) => sheetOf(name),
        getItemOrNullObject: (name: string) => ({ load: () => undefined, isNullObject: !sheets[name] }),
        getActiveWorksheet: () => sheetOf(Object.keys(sheets)[0]),
      },
    },
    sync: async () => undefined,
  };
  return ctx as unknown as Excel.RequestContext;
}

const setFormula = (sheetName: string, address: string, formula: string) =>
  ({ type: 'SET_FORMULA', sheetName, address, formula }) as RichAction;
const setCell = (sheetName: string, address: string, value: unknown) =>
  ({ type: 'SET_CELL', sheetName, address, value }) as RichAction;
const batch = (sheetName: string, operations: Array<{ address: string; value?: unknown; formula?: string }>) =>
  ({ type: 'BATCH_SET', sheetName, operations }) as RichAction;

async function guardError(promise: Promise<unknown>) {
  try {
    await promise;
    return null;
  } catch (error) {
    return error;
  }
}

describe('preflightOverwriteGuard — a block leaves nothing written (TASKS.md #319)', () => {
  it('rejects the whole card when a LATER action targets an occupied cell', async () => {
    const ctx = makeWorkbook({ Main: { A19: { value: '#SPILL!', formula: '=LET(x,1,x)' } } });
    const card = [
      batch('Main', [{ address: 'A1', value: 'Payments Dashboard' }]),
      { type: 'FORMAT_RANGE', sheetName: 'Main', range: 'A1' } as RichAction,
      setFormula('Main', 'A19', '=VSTACK(January!A2:M500)'),
    ];

    const error = await guardError(preflightOverwriteGuard(card, ctx));
    expect(isOverwriteGuardError(error)).toBe(true);
  });

  it('passes a card whose targets are empty', async () => {
    const ctx = makeWorkbook({ Main: {} });
    await expect(
      preflightOverwriteGuard([setCell('Main', 'A1', 'Title'), setFormula('Main', 'B2', '=1+1')], ctx),
    ).resolves.toBeUndefined();
  });

  it('skips a sheet the same card creates, and a sheet that does not exist yet', async () => {
    const ctx = makeWorkbook({ Main: {} });
    await expect(
      preflightOverwriteGuard(
        [
          { type: 'ADD_SHEET', name: 'January' } as unknown as RichAction,
          setCell('January', 'A1', 'Unit No'),
          setCell('Nowhere', 'A1', 'x'),
        ],
        ctx,
      ),
    ).resolves.toBeUndefined();
  });

  it('stops checking at a structural action — a later write may target cells it emptied', async () => {
    const ctx = makeWorkbook({ Main: { A1: { value: 'old' } } });
    await expect(
      preflightOverwriteGuard(
        [{ type: 'CLEAR_RANGE', sheetName: 'Main', range: 'A1' } as RichAction, setCell('Main', 'A1', 'new')],
        ctx,
      ),
    ).resolves.toBeUndefined();
  });
});

describe('idempotent writes — a cell already holding the intended content is not an overwrite (TASKS.md #319)', () => {
  it('lets the half-applied live step 6 retry: same title, same formula', async () => {
    const ctx = makeWorkbook({
      Main: {
        A1: { value: 'Payments Dashboard' },
        A19: { value: '#SPILL!', formula: '=LET(rows, VSTACK(January!A2:M500), rows)' },
      },
    });
    await expect(
      preflightOverwriteGuard(
        [
          batch('Main', [{ address: 'A1', value: 'Payments Dashboard' }]),
          setFormula('Main', 'A19', '=LET(rows,VSTACK(January!A2:M500),rows)'),
        ],
        ctx,
      ),
    ).resolves.toBeUndefined();
  });

  it('still blocks a DIFFERENT value or formula in the same cell', async () => {
    const ctx = makeWorkbook({ Main: { A1: { value: 'Payments Dashboard' }, B2: { value: 5, formula: '=2+3' } } });
    expect(isOverwriteGuardError(await guardError(guardAgainstOverwrite(setCell('Main', 'A1', 'Other'), ctx)))).toBe(true);
    expect(
      isOverwriteGuardError(await guardError(guardAgainstOverwrite(setFormula('Main', 'B2', '=9+9'), ctx))),
    ).toBe(true);
  });

  it('treats a formula returning "" as content, compared by formula text', async () => {
    const ctx = makeWorkbook({ Main: { G19: { value: '', formula: '=IF(E19="","",F19-E19)' } } });
    await expect(
      guardAgainstOverwrite(setFormula('Main', 'G19', '=IF(E19="","",F19-E19)'), ctx),
    ).resolves.toBeUndefined();
    expect(
      isOverwriteGuardError(await guardError(guardAgainstOverwrite(setFormula('Main', 'G19', '=1'), ctx))),
    ).toBe(true);
  });

  it('never matches a value write against a formula cell that happens to display it', () => {
    expect(cellAlreadyHolds({ value: 'Paid', formula: '=A1' }, { value: 'Paid' })).toBe(false);
    expect(cellAlreadyHolds({ value: 101, formula: '101' }, { value: '101' })).toBe(true);
  });
});
