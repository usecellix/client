import { readFileSync } from 'fs';
import { join } from 'path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  contiguousRuns,
  gridToRects,
  restoreRowSnapshot,
  sessionUndoStore,
  SESSION_UNDOABLE_ACTION_TYPES,
  withoutSessionUndoable,
  type RowSnapshot,
} from './rowSnapshot';

/* global Excel */

const ops: string[] = [];

function makeCtx() {
  const sheet = {
    getRangeByIndexes: vi.fn((row: number, col: number, rowCount: number, colCount: number) => {
      const range = {
        getEntireRow: () => ({
          insert: (dir: unknown) => ops.push(`insert@${row}:${String(dir)}`),
        }),
        set numberFormat(v: string[][]) {
          ops.push(`nf@${row}`);
          void v;
        },
        set formulas(v: unknown[][]) {
          ops.push(`formulas@${row},${col}:${JSON.stringify(v[0])}`);
        },
        format: {
          font: {
            set bold(v: boolean) {
              ops.push(`bold@${row}x${col}:${rowCount}x${colCount}=${v}`);
            },
          },
        },
      };
      return range;
    }),
  };
  return {
    workbook: {
      worksheets: {
        getItem: () => sheet,
        getActiveWorksheet: () => sheet,
        getItemOrNullObject: () => ({ isNullObject: false, load: () => undefined, ...sheet }),
      },
    },
    sync: vi.fn(async () => undefined),
  };
}

const snapshot = (over: Partial<RowSnapshot> = {}): RowSnapshot => ({
  kind: 'deleted-rows',
  sheetName: 'T_Delete',
  restorable: true,
  colStart: 0,
  colCount: 2,
  rows: [
    { index: 5, formulas: ['I-6', 300], numberFormats: ['General', '#,##0.00'] },
    { index: 1, formulas: ['I-2', 100], numberFormats: ['General', '#,##0.00'] },
    { index: 3, formulas: ['I-4', '=C4*2'], numberFormats: ['General', '#,##0.00'] },
  ],
  formats: [],
  ...over,
});

beforeEach(() => {
  ops.length = 0;
  (globalThis as unknown as { Excel: unknown }).Excel = { InsertShiftDirection: { down: 'Down' } };
});
afterEach(() => {
  sessionUndoStore.clear();
  delete (globalThis as unknown as { Excel?: unknown }).Excel;
});

describe('P1.6 row snapshot helpers', () => {
  it('contiguousRuns groups ascending runs', () => {
    expect(contiguousRuns([9, 3, 4, 5])).toEqual([
      [3, 5],
      [9, 9],
    ]);
    expect(contiguousRuns([])).toEqual([]);
  });

  it('gridToRects merges equal neighbours into rectangles', () => {
    expect(gridToRects([[0, 0, 1], [0, 0, 1]])).toEqual([
      { row: 0, col: 0, rowCount: 2, colCount: 2, index: 0 },
      { row: 0, col: 2, rowCount: 2, colCount: 1, index: 1 },
    ]);
  });

  it('strips the session-undoable types from an irreversible list', () => {
    expect(withoutSessionUndoable(['DELETE_MATCHING_ROWS', 'FREEZE_PANES', 'SET_MATCHING_ROWS'])).toEqual(['FREEZE_PANES']);
    expect(withoutSessionUndoable(undefined)).toEqual([]);
    expect([...SESSION_UNDOABLE_ACTION_TYPES]).toEqual(['DELETE_MATCHING_ROWS', 'SET_MATCHING_ROWS', 'DEFINE_NAMED_RANGE']);
  });
});

describe('P1.6 restore', () => {
  it('re-inserts deleted rows in ASCENDING order so each lands at its original position', async () => {
    const ctx = makeCtx();
    const report = await restoreRowSnapshot(snapshot(), ctx as never);
    const inserts = ops.filter((o) => o.startsWith('insert@')).map((o) => Number(o.split('@')[1].split(':')[0]));
    expect(inserts).toEqual([1, 3, 5]);
    expect(report.rowsRestored).toBe(3);
  });

  it('writes the original formulas/constants back into each restored row', async () => {
    const ctx = makeCtx();
    await restoreRowSnapshot(snapshot(), ctx as never);
    expect(ops).toContain('formulas@1,0:["I-2",100]');
    expect(ops).toContain('formulas@3,0:["I-4","=C4*2"]');
    expect(ops).toContain('formulas@5,0:["I-6",300]');
    // inserts all happen before any content write
    const lastInsert = Math.max(...ops.map((o, i) => (o.startsWith('insert@') ? i : -1)));
    const firstWrite = ops.findIndex((o) => o.startsWith('formulas@'));
    expect(lastInsert).toBeLessThan(firstWrite);
  });

  it('set-cells restore writes back only the one cell and inserts nothing', async () => {
    const ctx = makeCtx();
    const snap = snapshot({
      kind: 'set-cells',
      colStart: 3,
      colCount: 1,
      rows: [
        { index: 4, formulas: ['Open'], numberFormats: ['General'] },
        { index: 5, formulas: ['Open'], numberFormats: ['General'] },
      ],
    });
    await restoreRowSnapshot(snap, ctx as never);
    expect(ops.some((o) => o.startsWith('insert@'))).toBe(false);
    expect(ops).toContain('formulas@4,3:["Open"]');
    expect(ops).toContain('formulas@5,3:["Open"]');
  });

  it('applies the captured row formatting after the rows are back', async () => {
    const ctx = makeCtx();
    const snap = snapshot({
      formats: [
        {
          kind: 'format',
          sheetName: 'T_Delete',
          restorable: true,
          row: 1,
          col: 0,
          rowCount: 1,
          colCount: 2,
          palette: [{ bold: true }],
          grid: [[0, 0]],
        },
      ],
    });
    await restoreRowSnapshot(snap, ctx as never);
    expect(ops).toContain('bold@1x0:1x2=true');
  });

  it('counts, but does not fail on, a row-format run that could not be read at capture time', async () => {
    const ctx = makeCtx();
    const snap = snapshot({
      formats: [{ kind: 'format', sheetName: 'T_Delete', restorable: false, reason: 'borders' }],
    });
    const report = await restoreRowSnapshot(snap, ctx as never);
    expect(report.formatRunsSkipped).toBe(1);
    expect(report.rowsRestored).toBe(3);
  });

  it('refuses a snapshot that was not restorable', async () => {
    const ctx = makeCtx();
    await expect(
      restoreRowSnapshot({ ...snapshot(), restorable: false, reason: 'more than 2000 rows match' }, ctx as never),
    ).rejects.toThrow(/more than 2000 rows/);
  });
});

describe('P1.6 session store', () => {
  it('registers, reports, and forgets', () => {
    sessionUndoStore.register('cs-1', [snapshot()]);
    expect(sessionUndoStore.has('cs-1')).toBe(true);
    expect(sessionUndoStore.isRestorable('cs-1')).toBe(true);
    sessionUndoStore.register('cs-2', [snapshot({ restorable: false, reason: 'x' })]);
    expect(sessionUndoStore.isRestorable('cs-2')).toBe(false);
    sessionUndoStore.register('cs-3', undefined);
    expect(sessionUndoStore.has('cs-3')).toBe(false);
    sessionUndoStore.clear('cs-1');
    expect(sessionUndoStore.has('cs-1')).toBe(false);
  });

  it('never touches storage, the network or the console (module memory only)', () => {
    const source = readFileSync(join(__dirname, 'rowSnapshot.ts'), 'utf8')
      .split(/\r?\n/)
      .filter((line) => !line.trim().startsWith('*') && !line.trim().startsWith('//') && !line.trim().startsWith('/*'))
      .join(' ');
    expect(source).not.toMatch(/localStorage|sessionStorage|indexedDB|fetch\(|XMLHttpRequest|console\./);
  });
});

describe('P3.15 named-range undo', () => {
  function namesCtx(existing: Array<{ name: string; formula: string }>) {
    const log: string[] = [];
    const items = existing.map((n) => ({ ...n, delete: () => log.push(`delete ${n.name}`) }));
    return {
      log,
      ctx: {
        workbook: { names: { items, load: () => undefined, add: (n: string, f: string) => log.push(`add ${n}=${f}`) } },
        sync: async () => undefined,
      },
    };
  }
  const snap = (previousFormula: string | null): RowSnapshot => ({
    kind: 'named-range', sheetName: '', restorable: true, colStart: 0, colCount: 0, rows: [], formats: [],
    rangeName: 'Amounts', previousFormula,
  });

  it('a name that did not exist before is deleted on undo', async () => {
    const { ctx, log } = namesCtx([{ name: 'Amounts', formula: '=T_Days!$D$2:$D$6' }]);
    await restoreRowSnapshot(snap(null), ctx as never);
    expect(log).toEqual(['delete Amounts']);
  });

  it('a name that existed is put back to its old range', async () => {
    const { ctx, log } = namesCtx([{ name: 'Amounts', formula: '=T_Days!$D$2:$D$6' }]);
    await restoreRowSnapshot(snap('=Sheet1!$A$1:$A$3'), ctx as never);
    expect(log).toEqual(['delete Amounts', 'add Amounts==Sheet1!$A$1:$A$3']);
  });
});
