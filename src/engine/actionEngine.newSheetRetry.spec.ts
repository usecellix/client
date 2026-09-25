import { afterEach, describe, expect, it, vi } from 'vitest';
import { RichActionEngine } from './actionEngine';
import { OverwriteGuardError } from './overwriteGuard';
import type { RichAction } from '@/action.types';

/**
 * TASKS.md #326 — live run_1790330217489_abx5yzj: ADD_SHEET Main (position 0)
 * then SET_CELL Main!A1. The write's guard read of Main threw ItemNotFound,
 * while the FORMAT_RANGE on Main right after it succeeded — the sheet existed,
 * Excel had not settled it yet, and the dashboard lost its title.
 */
function itemNotFound() {
  return Object.assign(new Error("The requested resource doesn't exist."), {
    name: 'RichApi.Error',
    code: 'ItemNotFound',
  });
}

function stubExcel() {
  vi.stubGlobal('Excel', {
    run: async (fn: (ctx: unknown) => Promise<void>) => fn({ sync: async () => undefined }),
  });
}

const addMain = { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main', position: 0 } as RichAction;
const titleOn = (sheetName: string) =>
  ({ type: 'SET_CELL', sheetName, address: 'A1', value: 'Payments Dashboard' }) as RichAction;

type Dispatch = (action: RichAction) => Promise<unknown>;
function engineWith(dispatch: Dispatch) {
  const engine = new RichActionEngine();
  const spy = vi
    .spyOn(engine as unknown as { dispatch: Dispatch }, 'dispatch')
    .mockImplementation(dispatch);
  return { engine, spy };
}

describe('RichActionEngine — a just-created sheet not yet settled (TASKS.md #326)', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('retries once, through the full guarded dispatch, and the write lands', async () => {
    stubExcel();
    let titleAttempts = 0;
    const { engine, spy } = engineWith(async (action) => {
      if (action.type === 'ADD_SHEET') return { requestedName: 'Main', actualName: 'Main', reusedExisting: false };
      titleAttempts += 1;
      if (titleAttempts === 1) throw itemNotFound();
      return undefined;
    });

    const result = await engine.applyActions([addMain, titleOn('Main')]);

    expect(result.errors).toEqual([]);
    expect(result.applied).toBe(2);
    expect(spy).toHaveBeenCalledTimes(3);
  });

  it('does not retry ItemNotFound on a sheet this card did not create — that is a real failure', async () => {
    stubExcel();
    const { engine, spy } = engineWith(async () => {
      throw itemNotFound();
    });

    const result = await engine.applyActions([titleOn('Elsewhere')]);

    expect(spy).toHaveBeenCalledTimes(1);
    expect(result.errors[0]).toMatch(/SET_CELL: The requested resource doesn't exist/);
  });

  it('never retries an overwrite-guard block — it still stops the card unchanged', async () => {
    stubExcel();
    const block = new OverwriteGuardError({ message: 'Write blocked', targetRange: 'A1', sampleExistingValues: [] });
    const { engine, spy } = engineWith(async (action) => {
      if (action.type === 'ADD_SHEET') return { requestedName: 'Main', actualName: 'Main', reusedExisting: false };
      throw block;
    });

    await expect(engine.applyActions([addMain, titleOn('Main')])).rejects.toBe(block);
    expect(spy).toHaveBeenCalledTimes(2);
  });
});
