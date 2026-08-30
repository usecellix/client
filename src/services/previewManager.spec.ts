import { describe, expect, it, vi, afterEach } from 'vitest';
import { PreviewManager } from './previewManager';
import { SheetAction } from '@/types/sheet-actions';

/* global globalThis */

/**
 * TASKS.md #148 — nothing writes to the workbook before Accept.
 *
 * `previewManager.ts` had no test file at all until now (a gap
 * `CODEBASE_ANALYSIS.md` §3.3's #146 update called out explicitly), which is
 * part of why its soft-preview write survived long enough to cause #145, #146
 * and #147: three Office.js failures whose common root was code touching a
 * sheet that existed only because the preview had partially run.
 *
 * The mock below is deliberately strict about the thing that actually matters:
 * a worksheet the batch is about to create does not exist yet, so ANY accessor
 * reaching for it throws exactly as the live host does.
 */

const MONTHS = ['January', 'February', 'March'];

const HEADERS = ['Unit No', 'Guest', 'Guest Name', 'Check In'];

function freshBuildBatch(): SheetAction[] {
  const actions: SheetAction[] = [];
  for (const month of MONTHS) actions.push({ type: 'ADD_SHEET', name: month, sheetName: month });
  for (const month of MONTHS) {
    HEADERS.forEach((header, col) =>
      actions.push({ type: 'SET_CELL', sheetName: month, row: 0, col, value: header }),
    );
  }
  return actions;
}

/** Only `existing` sheets resolve; everything else behaves like a missing sheet. */
function installExcelMock(existing: string[] = []) {
  const applied: SheetAction[] = [];

  const makeRange = () => ({ load: vi.fn(), values: [['old value']] });

  const getItemOrNullObject = vi.fn((name: string) => {
    const isPresent = existing.some((n) => n.toLowerCase() === name.toLowerCase());
    return {
      isNullObject: !isPresent,
      name,
      load: vi.fn(),
      getRangeByIndexes: () => {
        if (!isPresent) {
          throw new Error("RichApi.Error: The requested resource doesn't exist.");
        }
        return makeRange();
      },
      getRange: () => {
        if (!isPresent) {
          throw new Error("RichApi.Error: The requested resource doesn't exist.");
        }
        return makeRange();
      },
      getUsedRangeOrNullObject: () => ({
        isNullObject: !isPresent,
        rowCount: 1,
        columnCount: 1,
        rowIndex: 0,
        columnIndex: 0,
        values: [['old value']],
        load: vi.fn(),
      }),
    };
  });

  const getActiveWorksheet = vi.fn(() => ({
    isNullObject: false,
    name: 'Sheet1',
    load: vi.fn(),
    getRangeByIndexes: () => makeRange(),
    getRange: () => makeRange(),
  }));

  (globalThis as Record<string, unknown>).Excel = {
    run: async (cb: (ctx: unknown) => Promise<unknown>) =>
      cb({
        workbook: { worksheets: { getItemOrNullObject, getActiveWorksheet } },
        sync: vi.fn(async () => undefined),
      }),
  };

  return { applied, getItemOrNullObject, getActiveWorksheet };
}

describe('PreviewManager — no writes before Accept (TASKS.md #148)', () => {
  const originalExcel = (globalThis as Record<string, unknown>).Excel;

  afterEach(() => {
    (globalThis as Record<string, unknown>).Excel = originalExcel;
    vi.restoreAllMocks();
  });

  it('renders a diff for a fresh multi-sheet build without touching Excel for the new sheets', async () => {
    // The exact shape that produced "The requested resource doesn't exist."
    const { getItemOrNullObject } = installExcelMock([]);
    const pm = new PreviewManager();

    const diff = await pm.render({ actions: freshBuildBatch(), summary: 'build' });

    // Every created sheet is diffed predictively — no lookup is even attempted.
    expect(getItemOrNullObject).not.toHaveBeenCalled();
    expect(diff.length).toBe(MONTHS.length * HEADERS.length);
    expect(diff.every((d) => d.before === '(empty)')).toBe(true);
  });

  it('predicts the right addresses for a created sheet', async () => {
    installExcelMock([]);
    const pm = new PreviewManager();

    const diff = await pm.render({ actions: freshBuildBatch(), summary: 'build' });
    const january = diff.filter((d) => d.sheetName === 'January');

    expect(january.map((d) => d.address)).toEqual(['A1', 'B1', 'C1', 'D1']);
    expect(january.map((d) => d.after)).toEqual(HEADERS);
  });

  it('predicts row 0 for an ADD_ROW-shaped header on a created sheet', async () => {
    installExcelMock([]);
    const pm = new PreviewManager();

    const diff = await pm.render({
      actions: [
        { type: 'ADD_SHEET', name: 'November', sheetName: 'November' },
        { type: 'ADD_ROW', sheetName: 'November', data: [...HEADERS] },
      ],
      summary: 'build',
    });

    expect(diff).toHaveLength(1);
    expect(diff[0].address).toBe('A1:D1');
    expect(diff[0].before).toBe('(empty)');
  });

  it('still live-reads `before` for a sheet that already exists', async () => {
    const { getItemOrNullObject } = installExcelMock(['Ledger']);
    const pm = new PreviewManager();

    const diff = await pm.render({
      actions: [{ type: 'SET_CELL', sheetName: 'Ledger', row: 0, col: 0, value: 'new' }],
      summary: 'edit',
    });

    expect(getItemOrNullObject).toHaveBeenCalledWith('Ledger');
    expect(diff[0].before).toBe('old value');
  });

  it('reject is a genuine no-op — it never writes anything', async () => {
    installExcelMock([]);
    const pm = new PreviewManager();
    await pm.render({ actions: freshBuildBatch(), summary: 'build' });
    expect(pm.active).toBe(true);

    // Would throw if it tried to apply revert actions through the mock.
    await expect(pm.reject()).resolves.toBeUndefined();
    expect(pm.active).toBe(false);
  });

  it('accept applies the FULL batch — creates are never treated as already-done', async () => {
    installExcelMock([]);
    const pm = new PreviewManager();
    const batch = freshBuildBatch();
    await pm.render({ actions: batch, summary: 'build' });

    const applySpy = vi
      .spyOn(
        await import('@/utils/actionEngine').then((m) => m.ActionEngine),
        'applyActionsWithReport',
      )
      .mockResolvedValue({ applied: batch.length, errors: [] });

    await pm.accept();

    expect(applySpy).toHaveBeenCalledTimes(1);
    const toApply = applySpy.mock.calls[0][0];
    // The regression this guards: with a soft preview, ADD_SHEETs were
    // subtracted here as "already applied", so Accept wrote into sheets it
    // assumed existed. All of them must be present.
    expect(toApply.filter((a) => a.type === 'ADD_SHEET')).toHaveLength(MONTHS.length);
    expect(toApply).toHaveLength(batch.length);
  });
});
