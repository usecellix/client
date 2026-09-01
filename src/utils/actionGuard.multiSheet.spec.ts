import { describe, expect, it } from 'vitest';
import { computeSheetLayout, sanitizeActions } from '@/utils/actionGuard';
import type { SheetGuardStates } from '@/engine/sheetGuardState';
import { SheetAction } from '@/types/sheet-actions';

/**
 * TASKS.md #137 — the multi-sheet header-guard data loss.
 *
 * Repro: "multiple sheets for all months in a year, plus a Main sheet with a
 * dashboard". The planner produced 12 month sheets × 10 header cells in row 1,
 * plus Main's A1 label. `sanitizeActions` filtered on `row === 0` without ever
 * reading `sheetName`, merged all 121 into ONE `ADD_ROW` carrying no sheet
 * name, and `resolveWorksheet`'s missing-name branch appended it to whatever
 * tab was active. Result: 12 empty month sheets, and one mangled header row on
 * December whose column A held Main's "Dashboard" label.
 */

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const HEADERS = [
  'Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out',
  'Rate Per Night', 'Total Amount', 'Source', 'Payment Status', 'Bank Account',
];

/** The exact wave-2 shape from the incident: 121 row-0 writes across 13 sheets. */
function buildMonthlyLedgerBatch(): SheetAction[] {
  const actions: SheetAction[] = [];
  for (const month of MONTHS) {
    HEADERS.forEach((header, col) => {
      actions.push({ type: 'SET_CELL', sheetName: month, row: 0, col, value: header });
    });
  }
  // Main's dashboard label — the cell that won column A on December.
  actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'Dashboard' });
  // Main body content, all below row 1, which survived the original bug.
  actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 1, col: 0, value: 'Totals' });
  return actions;
}

/** All 13 sheets freshly created and empty — what the live probe reports. */
function emptySheetStates(names: string[]): SheetGuardStates {
  return new Map(names.map((n) => [n.toLowerCase(), { hasHeaderRow: false, exists: true }]));
}

describe('sanitizeActions — multi-sheet header writes (TASKS.md #137)', () => {
  const allSheets = [...MONTHS, 'Main'];

  it('keeps every header write when the target sheets are empty', () => {
    const batch = buildMonthlyLedgerBatch();
    const result = sanitizeActions(batch, undefined, {
      sheetStates: emptySheetStates(allSheets),
    });

    // The regression: 122 in, 69 out. Every action must survive.
    expect(result.actions).toHaveLength(batch.length);
    expect(result.blocked).toHaveLength(0);
    expect(result.actions.some((a) => a.type === 'ADD_ROW')).toBe(false);
  });

  it('preserves each sheet its own complete header row', () => {
    const result = sanitizeActions(buildMonthlyLedgerBatch(), undefined, {
      sheetStates: emptySheetStates(allSheets),
    });

    for (const month of MONTHS) {
      const forMonth = result.actions.filter((a) => a.sheetName === month && a.row === 0);
      expect(forMonth).toHaveLength(HEADERS.length);
      expect(forMonth.map((a) => a.value)).toEqual(HEADERS);
    }
  });

  it('never lets one sheet\'s row-1 label leak into another sheet', () => {
    const result = sanitizeActions(buildMonthlyLedgerBatch(), undefined, {
      sheetStates: emptySheetStates(allSheets),
    });

    // "Dashboard" belongs to Main and nowhere else. In the incident it landed
    // in column A of December.
    const dashboardWrites = result.actions.filter((a) => a.value === 'Dashboard');
    expect(dashboardWrites).toHaveLength(1);
    expect(dashboardWrites[0]?.sheetName).toBe('Main');

    const december = result.actions.filter((a) => a.sheetName === 'December');
    expect(december.every((a) => a.value !== 'Dashboard')).toBe(true);
    expect(december.find((a) => a.col === 0)?.value).toBe('Unit No');
  });

  it('groups header logic per sheet — a populated sheet does not condemn the empty ones', () => {
    const batch: SheetAction[] = [
      // 'Existing' already has headers: this write must be converted.
      { type: 'SET_CELL', sheetName: 'Existing', row: 0, col: 0, value: 'Overwrite' },
      // 'January' is brand new: this write must pass through untouched.
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0, value: 'Unit No' },
    ];

    const sheetStates: SheetGuardStates = new Map([
      ['existing', { hasHeaderRow: true, exists: true }],
      ['january', { hasHeaderRow: false, exists: true }],
    ]);

    const result = sanitizeActions(batch, undefined, { sheetStates });

    const januaryWrite = result.actions.find((a) => a.sheetName === 'January');
    expect(januaryWrite?.type).toBe('SET_CELL');
    expect(januaryWrite?.row).toBe(0);

    const converted = result.actions.find((a) => a.type === 'ADD_ROW');
    expect(converted).toBeDefined();
    // The merged action must carry its sheet — this is what stopped it landing
    // on the active tab.
    expect(converted?.sheetName).toBe('Existing');
  });

  it('stamps sheetName on the merged ADD_ROW so it can never fall back to the active sheet', () => {
    const result = sanitizeActions(
      [
        { type: 'SET_CELL', sheetName: 'Purchase Register', row: 0, col: 0, value: 'A' },
        { type: 'SET_CELL', sheetName: 'Purchase Register', row: 0, col: 1, value: 'B' },
      ],
      undefined,
      { sheetStates: new Map([['purchase register', { hasHeaderRow: true, exists: true }]]) },
    );

    const addRow = result.actions.find((a) => a.type === 'ADD_ROW');
    expect(addRow?.sheetName).toBe('Purchase Register');
    expect(addRow?.data).toEqual(['A', 'B']);
  });

  it('reports what it dropped instead of shrinking the batch silently', () => {
    const result = sanitizeActions(
      [{ type: 'DELETE_ROW', sheetName: 'Ledger', row: 0 }],
      undefined,
      { sheetStates: new Map([['ledger', { hasHeaderRow: true, exists: true }]]) },
    );

    expect(result.actions).toHaveLength(0);
    expect(result.blocked).toHaveLength(1);
    expect(result.warnings.join(' ')).toContain('Blocked 1 action');
  });

  it('no longer depends on an unrelated ADD_ROW being present in the batch', () => {
    // The old `detectPopulateEmptySheet` required BOTH a row-0 SET_CELL and an
    // ADD_ROW to conclude "empty". Whether an ADD_ROW happened to be in the
    // batch therefore decided if headers survived — the same prompt could work
    // or fail run to run. Live sheet state must give the same answer either way.
    const headerOnly: SheetAction[] = [
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0, value: 'Unit No' },
    ];
    const withAddRow: SheetAction[] = [
      ...headerOnly,
      { type: 'ADD_ROW', sheetName: 'January', data: ['x'] },
    ];
    const sheetStates: SheetGuardStates = new Map([
      ['january', { hasHeaderRow: false, exists: true }],
    ]);

    const a = sanitizeActions(headerOnly, undefined, { sheetStates });
    const b = sanitizeActions(withAddRow, undefined, { sheetStates });

    expect(a.actions.filter((x) => x.type === 'SET_CELL')).toHaveLength(1);
    expect(b.actions.filter((x) => x.type === 'SET_CELL')).toHaveLength(1);
    expect(a.blocked).toHaveLength(0);
    expect(b.blocked).toHaveLength(0);
  });

  it('still protects a real single-sheet header row (original guard intent)', () => {
    // Unchanged behaviour from actionGuard.spec.ts: no probe result, active
    // sheet has headers, so a row-0 data write is converted, not applied.
    const layout = computeSheetLayout([
      ['Date', 'Supplier', 'Amount'],
      ['2024-01-01', 'Acme', 100],
    ]);

    const result = sanitizeActions(
      [
        { type: 'SET_CELL', row: 0, col: 0, value: 'New Date label' },
        { type: 'SET_CELL', row: 0, col: 1, value: 'New Supplier' },
      ],
      layout,
    );

    expect(result.actions.some((a) => a.type === 'ADD_ROW')).toBe(true);
    expect(result.actions.every((a) => a.row !== 0 || a.type === 'ADD_ROW')).toBe(true);
  });
});
