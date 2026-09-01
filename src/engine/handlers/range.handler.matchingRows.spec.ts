import { describe, expect, it, vi } from 'vitest';
import type { SetMatchingRowsAction } from '@/action.types';
import { handleSetMatchingRows } from './range.handler';

/* global Excel */

const HEADERS = ['Date', 'Invoice No', 'Payment Status', 'Remarks'];

/** 50 data rows alternating Paid/Pending, like the Purchase Register under test. */
function buildRegisterRows(dataRowCount: number): unknown[][] {
  const rows: unknown[][] = [HEADERS];
  for (let i = 0; i < dataRowCount; i++) {
    rows.push([
      `2026-01-${String((i % 28) + 1).padStart(2, '0')}`,
      `INV-${1001 + i}`,
      i % 3 === 1 ? 'Pending' : 'Paid',
      '',
    ]);
  }
  return rows;
}

/**
 * Mock worksheet whose used range is the FULL sheet, while the action asks for a
 * truncated range — the F13 shape (executor sampled ~10 rows of 50).
 */
function makeCtx(totalDataRows: number) {
  const allRows = buildRegisterRows(totalDataRows);
  const writes: Array<{ row: number; col: number; value: unknown }> = [];

  const usedRange = { rowIndex: 0, rowCount: allRows.length, load: vi.fn() };

  // getRangeByIndexes serves both the widened scan range and the per-cell writes.
  const getRangeByIndexes = vi.fn(
    (row: number, col: number, rowCount: number, colCount: number) => {
      if (rowCount === 1 && colCount === 1) {
        const cell = {
          set values(v: unknown[][]) {
            writes.push({ row, col, value: v[0][0] });
          },
          get values() {
            return [[null]];
          },
        };
        return cell as unknown as Excel.Range;
      }
      return {
        values: allRows.slice(row, row + rowCount).map((r) => r.slice(col, col + colCount)),
        rowIndex: row,
        columnIndex: col,
        columnCount: colCount,
        load: vi.fn(),
      } as unknown as Excel.Range;
    },
  );

  const getRange = vi.fn((address: string) => {
    // Narrow range the model guessed — only the sampled rows.
    const match = /^[A-Z]+(\d+):[A-Z]+(\d+)$/.exec(address);
    const startRow = match ? Number(match[1]) - 1 : 0;
    const endRow = match ? Number(match[2]) - 1 : allRows.length - 1;
    return {
      values: allRows.slice(startRow, endRow + 1),
      rowIndex: startRow,
      columnIndex: 0,
      columnCount: HEADERS.length,
      load: vi.fn(),
    } as unknown as Excel.Range;
  });

  const sheet = {
    getUsedRange: vi.fn(() => usedRange),
    getRange,
    getRangeByIndexes,
  };

  const ctx = {
    workbook: {
      worksheets: {
        getItem: vi.fn(() => sheet),
        getActiveWorksheet: vi.fn(() => sheet),
      },
    },
    sync: vi.fn(async () => undefined),
  } as unknown as Excel.RequestContext;

  return { ctx, writes, sheet };
}

describe('handleSetMatchingRows — truncated executor range (TASKS.md F13)', () => {
  it('updates every matching row on the sheet, not just the sampled window', async () => {
    // Real incident: 50-row register, executor emitted A1:L11, only 10 cells written.
    const { ctx, writes } = makeCtx(50);
    const action: SetMatchingRowsAction = {
      type: 'SET_MATCHING_ROWS',
      sheetName: 'Purchase Register',
      range: 'A1:D11',
      hasHeaders: true,
      filter: { column: 'Payment Status', operator: 'equals', value: 'Pending' },
      targetColumn: 'Remarks',
      value: 'Follow up',
    };

    const result = await handleSetMatchingRows(action, ctx);

    // i % 3 === 1 over 50 rows => 17 Pending rows.
    expect(result.rowsUpdated).toBe(17);
    expect(writes).toHaveLength(17);
    // The last Pending row is far beyond the sampled window — the bug's signature.
    expect(Math.max(...writes.map((w) => w.row))).toBeGreaterThan(11);
    expect(writes.every((w) => w.value === 'Follow up')).toBe(true);
  });

  it('writes into the resolved target column', async () => {
    const { ctx, writes } = makeCtx(50);
    await handleSetMatchingRows(
      {
        type: 'SET_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A1:D11',
        hasHeaders: true,
        filter: { column: 'Payment Status', operator: 'equals', value: 'Paid' },
        targetColumn: 'Remarks',
        value: 'Cleared',
      },
      ctx,
    );
    expect(writes.every((w) => w.col === HEADERS.indexOf('Remarks'))).toBe(true);
  });

  it('does not shrink a range that already covers the used rows', async () => {
    const { ctx, writes } = makeCtx(12);
    const result = await handleSetMatchingRows(
      {
        type: 'SET_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A1:D13',
        hasHeaders: true,
        filter: { column: 'Payment Status', operator: 'equals', value: 'Pending' },
        targetColumn: 'Remarks',
        value: 'Follow up',
      },
      ctx,
    );
    expect(result.rowsUpdated).toBe(writes.length);
    expect(result.rowsUpdated).toBeGreaterThan(0);
  });

  it('reports rows scanned alongside rows updated (TASKS.md #93)', async () => {
    // A partial write used to look identical to a complete one. Coverage is now
    // reported so "17 of 50 rows" can be distinguished from "17 of 17".
    const { ctx } = makeCtx(50);
    const result = await handleSetMatchingRows(
      {
        type: 'SET_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A1:D11',
        hasHeaders: true,
        filter: { column: 'Payment Status', operator: 'equals', value: 'Pending' },
        targetColumn: 'Remarks',
        value: 'Follow up',
      },
      ctx,
    );
    expect(result.rowsScanned).toBe(50);
    expect(result.rowsUpdated).toBe(17);
  });

  it('fills the whole column when no filter is supplied', async () => {
    const { ctx, writes } = makeCtx(50);
    const result = await handleSetMatchingRows(
      {
        type: 'SET_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A1:D11',
        hasHeaders: true,
        targetColumn: 'Remarks',
        value: '',
      },
      ctx,
    );
    expect(result.rowsUpdated).toBe(50);
    expect(writes).toHaveLength(50);
  });
});
