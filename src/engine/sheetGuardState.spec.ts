import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import {
  probeSheetGuardStates,
  probeSheetGuardStatesSafe,
  sheetsCreatedInBatch,
  sheetKeyOf,
  targetSheetNames,
} from './sheetGuardState';
import { SheetAction } from '@/types/sheet-actions';

/* global globalThis */

interface FakeSheetSpec {
  /** Row 1 values, or `null` when the sheet does not exist. */
  headerRow: unknown[] | null;
}

function installExcelMock(sheets: Record<string, FakeSheetSpec>, activeHeaderRow: unknown[] = []) {
  const getItemOrNullObject = vi.fn((name: string) => {
    const spec = sheets[name];
    const sheet = {
      isNullObject: !spec,
      getRangeByIndexes: () => ({
        values: spec?.headerRow ? [spec.headerRow] : [[]],
        load: vi.fn(),
      }),
      load: vi.fn(),
    };
    return sheet;
  });

  const getActiveWorksheet = vi.fn(() => ({
    isNullObject: false,
    getRangeByIndexes: () => ({ values: [activeHeaderRow], load: vi.fn() }),
    load: vi.fn(),
  }));

  (globalThis as Record<string, unknown>).Excel = {
    run: async (cb: (ctx: unknown) => Promise<void>) =>
      cb({
        workbook: { worksheets: { getItemOrNullObject, getActiveWorksheet } },
        sync: vi.fn(async () => undefined),
      }),
  };

  return { getItemOrNullObject, getActiveWorksheet };
}

describe('sheetGuardState helpers', () => {
  it('keys actions by sheet, case-insensitively', () => {
    expect(sheetKeyOf({ type: 'SET_CELL', sheetName: 'January' })).toBe('january');
    expect(sheetKeyOf({ type: 'SET_CELL' })).toBe('');
  });

  it('lists distinct target sheets preserving original casing', () => {
    const actions: SheetAction[] = [
      { type: 'SET_CELL', sheetName: 'Main' },
      { type: 'SET_CELL', sheetName: 'Main' },
      { type: 'SET_CELL', sheetName: 'January' },
      { type: 'SET_CELL' },
    ];
    expect(targetSheetNames(actions)).toEqual(['Main', 'January']);
  });

  it('treats sheets created in the same batch as empty by construction', () => {
    const created = sheetsCreatedInBatch([
      { type: 'ADD_SHEET', name: 'January' },
      { type: 'CREATE_SHEET', sheetName: 'Main' },
      { type: 'SET_CELL', sheetName: 'February' },
    ]);
    expect(created).toEqual(new Set(['january', 'main']));
  });
});

describe('probeSheetGuardStates', () => {
  const originalExcel = (globalThis as Record<string, unknown>).Excel;

  afterEach(() => {
    (globalThis as Record<string, unknown>).Excel = originalExcel;
  });

  it('reports an empty row 1 as free to write', async () => {
    installExcelMock({ January: { headerRow: [] } });
    const states = await probeSheetGuardStates([
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0, value: 'Unit No' },
    ]);
    expect(states.get('january')).toEqual({ hasHeaderRow: false, exists: true });
  });

  it('reports an occupied row 1 as a real header row', async () => {
    installExcelMock({ Ledger: { headerRow: ['Date', 'Amount'] } });
    const states = await probeSheetGuardStates([
      { type: 'SET_CELL', sheetName: 'Ledger', row: 0, col: 0, value: 'x' },
    ]);
    expect(states.get('ledger')).toEqual({ hasHeaderRow: true, exists: true });
  });

  it('treats a not-yet-created sheet as empty rather than failing', async () => {
    installExcelMock({});
    const states = await probeSheetGuardStates([
      { type: 'SET_CELL', sheetName: 'December', row: 0, col: 0, value: 'x' },
    ]);
    expect(states.get('december')).toEqual({ hasHeaderRow: false, exists: false });
  });

  it('never queues a lookup for a sheet the batch is about to create', async () => {
    // The poisoned-context failure mode documented in worksheet.handler.ts:
    // a getItem() for a missing sheet fails the whole batch's next sync().
    const { getItemOrNullObject } = installExcelMock({});
    await probeSheetGuardStates([
      { type: 'ADD_SHEET', name: 'January' },
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0, value: 'Unit No' },
    ]);
    expect(getItemOrNullObject).not.toHaveBeenCalled();
  });

  it('probes the active sheet only when an action omits sheetName', async () => {
    const { getActiveWorksheet } = installExcelMock({ Main: { headerRow: [] } });
    await probeSheetGuardStates([{ type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0 }]);
    expect(getActiveWorksheet).not.toHaveBeenCalled();

    const second = installExcelMock({}, ['Date']);
    const states = await probeSheetGuardStates([{ type: 'SET_CELL', row: 0, col: 0 }]);
    expect(second.getActiveWorksheet).toHaveBeenCalled();
    expect(states.get('')).toEqual({ hasHeaderRow: true, exists: true });
  });
});

/**
 * TASKS.md #145 — a live Excel session failed with "The requested resource
 * doesn't exist." before a single sheet was created.
 *
 * `probeSheetGuardStates` chained `.getRangeByIndexes()` off a
 * `getItemOrNullObject()` result and loaded BOTH in the same `ctx.sync()` —
 * for a sheet that doesn't exist yet (the normal case for a fresh multi-sheet
 * build), Office.js resolves the range request against the still-unconfirmed
 * parent and throws, failing the WHOLE sync — the exact failure mode
 * `worksheet.handler.ts` was already fixed for once, reintroduced here.
 *
 * `installExcelMock` in the tests above never modeled this — it returns
 * `values` synchronously regardless of order, so it could not have caught
 * the bug. This mock enforces the real Office.js contract: any
 * `getRangeByIndexes()` call on a sheet whose `isNullObject` has not yet been
 * resolved by a PRIOR `ctx.sync()` throws, exactly as the live host does.
 */
function installStrictTwoPhaseMock(existingSheets: Set<string>) {
  let syncCount = 0;
  const resolvedNullObject = new Map<string, boolean>();
  const rangeRequests: { sheet: string; requestedAtSync: number }[] = [];

  const makeSheet = (name: string, isActive = false) => {
    let nullObjectLoaded = false;
    return {
      get isNullObject() {
        if (!nullObjectLoaded) {
          throw new Error(
            `PropertyNotLoaded: isNullObject read on "${name}" before load()+sync()`,
          );
        }
        return !existingSheets.has(name.toLowerCase());
      },
      load: (prop: string) => {
        if (prop === 'isNullObject') nullObjectLoaded = true;
      },
      getRangeByIndexes: () => {
        rangeRequests.push({ sheet: name, requestedAtSync: syncCount });
        const resolved = resolvedNullObject.get(name.toLowerCase());
        // The real-world failure: requesting a range whose parent sheet's
        // existence was never confirmed by a completed sync.
        if (resolved === undefined) {
          throw new Error(
            `RichApi.Error: The requested resource doesn't exist. (range on "${name}" requested before its sheet's existence was confirmed)`,
          );
        }
        if (resolved === false) {
          throw new Error(`RichApi.Error: The requested resource doesn't exist. (sheet "${name}" does not exist)`);
        }
        return {
          load: () => undefined,
          values: [['Unit No', 'Guest']],
        };
      },
      _markResolved: () => {
        resolvedNullObject.set(name.toLowerCase(), existingSheets.has(name.toLowerCase()));
      },
      _isActive: isActive,
    };
  };

  const sheetInstances = new Map<string, ReturnType<typeof makeSheet>>();

  const getItemOrNullObject = (name: string) => {
    const existing = sheetInstances.get(name.toLowerCase());
    if (existing) return existing;
    const created = makeSheet(name);
    sheetInstances.set(name.toLowerCase(), created);
    return created;
  };

  const getActiveWorksheet = () => {
    const created = makeSheet('__active__', true);
    sheetInstances.set('__active__', created);
    return created;
  };

  const sync = async () => {
    syncCount += 1;
    // A real sync resolves every `load()`ed property queued before it —
    // including isNullObject reads for sheets requested in this batch.
    for (const sheet of sheetInstances.values()) sheet._markResolved();
  };

  (globalThis as Record<string, unknown>).Excel = {
    run: async (cb: (ctx: unknown) => Promise<void>) =>
      cb({
        workbook: { worksheets: { getItemOrNullObject, getActiveWorksheet } },
        sync,
      }),
  };

  return { rangeRequests: () => rangeRequests, syncCount: () => syncCount };
}

describe('probeSheetGuardStates — two-phase sync (TASKS.md #145)', () => {
  const originalExcel = (globalThis as Record<string, unknown>).Excel;

  afterEach(() => {
    (globalThis as Record<string, unknown>).Excel = originalExcel;
  });

  it('never requests a range before that sheet\'s existence is confirmed by a sync', async () => {
    // The exact shape that failed live: 12 month sheets, none created yet.
    const mock = installStrictTwoPhaseMock(new Set());
    const actions: SheetAction[] = [
      'January', 'February', 'March', 'April', 'May', 'June',
    ].map((sheet) => ({ type: 'SET_CELL', sheetName: sheet, row: 0, col: 0 }));

    // Must not throw. Every sheet is missing, so no range should ever be
    // requested — that is the whole point of checking isNullObject first.
    const states = await probeSheetGuardStates(actions);

    expect(mock.rangeRequests()).toHaveLength(0);
    for (const sheet of ['January', 'February', 'March', 'April', 'May', 'June']) {
      expect(states.get(sheet.toLowerCase())).toEqual({ hasHeaderRow: false, exists: false });
    }
  });

  it('only requests ranges for sheets confirmed to exist, after a sync', async () => {
    const mock = installStrictTwoPhaseMock(new Set(['ledger']));
    const actions: SheetAction[] = [
      { type: 'SET_CELL', sheetName: 'Ledger', row: 0, col: 0 },
      { type: 'SET_CELL', sheetName: 'Missing', row: 0, col: 0 },
    ];

    const states = await probeSheetGuardStates(actions);

    const requests = mock.rangeRequests();
    expect(requests).toHaveLength(1);
    expect(requests[0].sheet).toBe('Ledger');
    // The range request happened strictly after the sync that resolved
    // isNullObject — never in the same batch.
    expect(requests[0].requestedAtSync).toBeGreaterThan(0);
    expect(states.get('missing')).toEqual({ hasHeaderRow: false, exists: false });
    expect(states.get('ledger')?.exists).toBe(true);
  });

  it('handles a full 12-sheet fresh-build probe without throwing', async () => {
    installStrictTwoPhaseMock(new Set());
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December', 'Main',
    ];
    const actions: SheetAction[] = months.map((sheet) => ({
      type: 'SET_CELL',
      sheetName: sheet,
      row: 0,
      col: 0,
    }));

    const states = await probeSheetGuardStates(actions);
    expect(states.size).toBe(months.length);
  });
});

describe('probeSheetGuardStatesSafe', () => {
  const originalExcel = (globalThis as Record<string, unknown>).Excel;

  beforeEach(() => {
    (globalThis as Record<string, unknown>).Excel = {
      run: async () => {
        throw new Error('Office.js unavailable');
      },
    };
  });

  afterEach(() => {
    (globalThis as Record<string, unknown>).Excel = originalExcel;
  });

  it('degrades to the legacy guard instead of blocking the apply', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    const states = await probeSheetGuardStatesSafe([
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0 },
    ]);
    expect(states.size).toBe(0);
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
  });
});
