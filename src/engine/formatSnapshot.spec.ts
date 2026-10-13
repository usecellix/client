import { afterEach, describe, expect, it, vi } from 'vitest';
import type { RichAction } from '@/action.types';
import { RichActionEngine } from './actionEngine';
import { captureFormatSnapshot, MAX_SNAPSHOT_CELLS, MAX_SNAPSHOT_COLUMNS } from './formatSnapshot';

/**
 * TASKS.md #400 — what a formatting action is about to overwrite is read off
 * Excel first, so Revert can put it back. These tests stand in for Excel with a
 * small fake whose cells hold known formats.
 */

const colIndex = (s: string) => s.toUpperCase().charCodeAt(0) - 65;

/** "A3:D3" / "B4" → 0-based origin and size (single letters are enough here). */
function bounds(address: string) {
  const m = /^([A-Za-z])(\d+)(?::([A-Za-z])(\d+))?$/.exec(address);
  if (!m) throw new Error(`fake cannot read ${address}`);
  const c1 = colIndex(m[1]);
  const r1 = Number(m[2]) - 1;
  const c2 = m[3] ? colIndex(m[3]) : c1;
  const r2 = m[4] ? Number(m[4]) - 1 : r1;
  return { rowIndex: r1, columnIndex: c1, rowCount: r2 - r1 + 1, columnCount: c2 - c1 + 1 };
}

interface FakeCell {
  bold?: boolean | null;
  italic?: boolean;
  size?: number;
  name?: string;
  color?: string;
  fillPattern?: string;
  fillColor?: string;
  tint?: number;
  h?: string;
  v?: string;
  wrap?: boolean;
  numberFormat?: string;
}

const plain: FakeCell = {
  bold: false,
  italic: false,
  size: 11,
  name: 'Aptos Narrow',
  color: '#000000',
  fillPattern: 'None',
  fillColor: '#FFFFFF',
  tint: 0,
  h: 'General',
  v: 'Bottom',
  wrap: false,
  numberFormat: 'General',
};

interface FakeSheet {
  /** Format of the cell at (row, col); anything not listed is `plain`. */
  cell: (row: number, col: number) => FakeCell;
  columnWidths?: Record<number, number>;
  usedColumns?: { columnIndex: number; columnCount: number } | null;
  /**
   * Whether this Excel reports a fill's pattern.
   *  - 'reported' (default): 'None' for no fill, 'Solid' for a solid one
   *  - 'null': null for every cell, filled or not. Excel on Windows, live, Oct 9 2026:
   *    a solid light-green cell read as { color: "#C6EFCE", pattern: null }.
   */
  fillPattern?: 'reported' | 'null';
  /** What a cell with no fill reads as for colour. Excel builds differ; default '#FFFFFF'. */
  unfilledColor?: string;
  /** The sheet's last cell, which the reader compares against. Unfilled unless a test says otherwise. */
  corner?: FakeCell;
}

function stubExcel(sheet: FakeSheet, extra: { sync?: () => Promise<void> } = {}) {
  const reportsPattern = (sheet.fillPattern ?? 'reported') === 'reported';
  /** One cell's fill, as this Excel would hand it back. */
  const rawFill = (c: FakeCell) => {
    const unfilled = c.fillPattern === 'None';
    return {
      color: unfilled ? (sheet.unfilledColor ?? c.fillColor) : c.fillColor,
      pattern: reportsPattern ? c.fillPattern : null,
      tintAndShade: c.tint,
    };
  };
  const rangeFor = (address: string) => {
    const b = bounds(address);
    const at = (r: number, c: number) => ({ ...plain, ...sheet.cell(b.rowIndex + r, b.columnIndex + c) });
    const grid = <T,>(pick: (cell: FakeCell) => T) =>
      Array.from({ length: b.rowCount }, (_, r) => Array.from({ length: b.columnCount }, (_, c) => pick(at(r, c))));
    // Like Excel: a property that was not loaded throws when read, it does not come back empty.
    const loaded = new Set<string>();
    return {
      ...b,
      address,
      load: vi.fn((names: string | string[]) => [names].flat().forEach((n) => loaded.add(n))),
      get numberFormat() {
        if (!loaded.has('numberFormat')) {
          throw new Error("The property 'numberFormat' is not available. Before reading the property's value, call the load method");
        }
        return grid((c) => c.numberFormat);
      },
      set numberFormat(_value: unknown) {
        // a write, e.g. by SET_CELL; the fake keeps no state for it
      },
      getCellProperties: vi.fn(() => ({
        value: grid((c) => ({
          format: {
            font: { bold: c.bold, italic: c.italic, size: c.size, name: c.name, color: c.color },
            fill: rawFill(c),
            horizontalAlignment: c.h,
            verticalAlignment: c.v,
            wrapText: c.wrap,
          },
        })),
      })),
      format: { font: {}, fill: { clear: vi.fn() }, borders: { getItem: vi.fn(() => ({})) } },
      select: vi.fn(),
    };
  };
  const worksheet = {
    getRange: vi.fn((address: string) => rangeFor(address)),
    getRangeByIndexes: vi.fn((_r: number, col: number) => ({
      format: {
        columnWidth: sheet.columnWidths?.[col] ?? 64,
        load: vi.fn(),
      },
      // Only the sheet's last cell is ever read this way: the reader's "no fill" reference.
      getCellProperties: vi.fn(() => ({ value: [[{ format: { fill: rawFill({ ...plain, ...sheet.corner }) } }]] })),
    })),
    getUsedRangeOrNullObject: vi.fn(() => ({
      isNullObject: sheet.usedColumns === null,
      columnIndex: sheet.usedColumns?.columnIndex ?? 0,
      columnCount: sheet.usedColumns?.columnCount ?? 0,
      load: vi.fn(),
    })),
    getUsedRange: vi.fn(() => ({ format: { autofitColumns: vi.fn() } })),
  };
  const ctx = {
    workbook: {
      worksheets: {
        getItem: vi.fn(() => worksheet),
        getItemOrNullObject: vi.fn(() => ({ isNullObject: false, load: vi.fn() })),
        getActiveWorksheet: vi.fn(() => worksheet),
      },
    },
    sync: vi.fn(extra.sync ?? (async () => undefined)),
  };
  const run = vi.fn(async (fn: (c: unknown) => Promise<void>) => fn(ctx));
  vi.stubGlobal('Excel', { run });
  return { ctx: ctx as never, worksheet };
}

const formatAction = (range: string, format: Record<string, unknown>) =>
  ({ type: 'FORMAT_RANGE', sheetName: 'Monthly Summary', range, format }) as unknown as RichAction;

afterEach(() => vi.unstubAllGlobals());

describe('captureFormatSnapshot — FORMAT_RANGE', () => {
  it('reads only the properties the action writes, one palette entry for cells that match', async () => {
    const { ctx } = stubExcel({ cell: () => ({}) });
    const snapshot = await captureFormatSnapshot(formatAction('A3:D3', { bold: true, fillColor: '#4472C4', fontColor: '#FFFFFF' }), ctx);

    expect(snapshot).toEqual({
      kind: 'format',
      sheetName: 'Monthly Summary',
      restorable: true,
      row: 2,
      col: 0,
      rowCount: 1,
      colCount: 4,
      palette: [{ bold: false, fontColor: '#000000', clearFill: true }],
      grid: [[0, 0, 0, 0]],
    });
  });

  it('keeps cells that differ apart, and the cell they came from', async () => {
    const { ctx } = stubExcel({ cell: (_r, c) => (c === 1 ? { bold: true, fillPattern: 'Solid', fillColor: '#FFFF00' } : {}) });
    const snapshot = await captureFormatSnapshot(formatAction('A3:C3', { bold: false, fillColor: '#FFFFFF' }), ctx);

    expect(snapshot).toMatchObject({ restorable: true, grid: [[0, 1, 0]] });
    expect(snapshot?.palette).toEqual([
      { bold: false, clearFill: true },
      { bold: true, fillColor: '#FFFF00' },
    ]);
  });

  it('reads number formats cell by cell, since a range can hold several', async () => {
    const { ctx } = stubExcel({ cell: (r) => (r === 3 ? { numberFormat: '0.00' } : { numberFormat: 'General' }) });
    const snapshot = await captureFormatSnapshot(formatAction('B4:D5', { numberFormat: '#,##0.00' }), ctx);

    expect(snapshot).toMatchObject({ row: 3, col: 1, rowCount: 2, colCount: 3, grid: [[0, 0, 0], [1, 1, 1]] });
    expect(snapshot?.palette).toEqual([{ numberFormat: '0.00' }, { numberFormat: 'General' }]);
  });

  it('turns Excel alignment names into the ones FORMAT_RANGE writes', async () => {
    const { ctx } = stubExcel({ cell: () => ({ h: 'Center', v: 'Center' }) });
    const snapshot = await captureFormatSnapshot(formatAction('A1', { horizontalAlignment: 'left', verticalAlignment: 'top', wrapText: true }), ctx);
    expect(snapshot?.palette).toEqual([{ horizontalAlignment: 'center', verticalAlignment: 'middle', wrapText: false }]);
  });

  it('treats empty font colour as black, which is what Excel draws', async () => {
    const { ctx } = stubExcel({ cell: () => ({ color: '' }) });
    const snapshot = await captureFormatSnapshot(formatAction('A1', { fontColor: '#FF0000' }), ctx);
    expect(snapshot?.palette).toEqual([{ fontColor: '#000000' }]);
  });

  it.each([
    ['a fill pattern FORMAT_RANGE cannot write', { fillPattern: 'Gray50' }, { fillColor: '#FF0000' }],
    ['a horizontal alignment it cannot write', { h: 'Justify' }, { horizontalAlignment: 'left' }],
    ['a vertical alignment it cannot write', { v: 'Distributed' }, { verticalAlignment: 'top' }],
    ['a value Excel would not return (mixed)', { bold: null }, { bold: true }],
  ])('refuses %s, never guesses', async (_label, cell, format) => {
    const { ctx } = stubExcel({ cell: () => cell });
    const snapshot = await captureFormatSnapshot(formatAction('A1:B2', format), ctx);
    expect(snapshot).toMatchObject({ kind: 'format', restorable: false });
    expect(snapshot?.reason).toBeTruthy();
    expect(snapshot?.grid).toBeUndefined();
  });

  // Live, Oct 9 2026, Excel on Windows: a cell with a solid light-green fill read as
  // { color: "#C6EFCE", pattern: null }, and so did the range. The pattern is how a fill is
  // told from no fill, so every fill change was refused. These are that Excel.
  describe('when Excel does not report the fill pattern at all', () => {
    it('reads a coloured cell as a solid fill of that colour (the live case)', async () => {
      const { ctx } = stubExcel({ fillPattern: 'null', cell: () => ({ fillPattern: 'Solid', fillColor: '#C6EFCE', bold: true }) });
      const snapshot = await captureFormatSnapshot(formatAction('A3:D3', { bold: true, fillColor: '#ADD8E6' }), ctx);

      expect(snapshot).toMatchObject({ restorable: true, grid: [[0, 0, 0, 0]] });
      expect(snapshot?.palette).toEqual([{ bold: true, fillColor: '#C6EFCE' }]);
    });

    it('reads a cell that looks exactly like the unformatted corner of the sheet as no fill', async () => {
      const { ctx } = stubExcel({ fillPattern: 'null', cell: () => ({}) });
      const snapshot = await captureFormatSnapshot(formatAction('A1:D1', { fillColor: '#FFF2CC' }), ctx);
      expect(snapshot?.palette).toEqual([{ clearFill: true }]);
    });

    it('reads an empty colour as no fill, for an Excel that reports it that way', async () => {
      const { ctx } = stubExcel({ fillPattern: 'null', unfilledColor: '', cell: () => ({}) });
      const snapshot = await captureFormatSnapshot(formatAction('A1:B1', { fillColor: '#FFF2CC' }), ctx);
      expect(snapshot?.palette).toEqual([{ clearFill: true }]);
    });

    it('keeps a white fill apart from no fill when Excel reports the two differently', async () => {
      // No fill reads as '', so a cell that reads '#FFFFFF' has a real white fill.
      const { ctx } = stubExcel({
        fillPattern: 'null',
        unfilledColor: '',
        cell: (_r, c) => (c === 1 ? { fillPattern: 'Solid', fillColor: '#FFFFFF' } : {}),
      });
      const snapshot = await captureFormatSnapshot(formatAction('A1:C1', { fillColor: '#FF0000' }), ctx);
      expect(snapshot?.palette).toEqual([{ clearFill: true }, { fillColor: '#FFFFFF' }]);
      expect(snapshot?.grid).toEqual([[0, 1, 0]]);
    });

    it('keeps filled and unfilled cells of one range apart', async () => {
      const { ctx } = stubExcel({
        fillPattern: 'null',
        cell: (_r, c) => (c === 1 ? { fillPattern: 'Solid', fillColor: '#FFFF00' } : {}),
      });
      const snapshot = await captureFormatSnapshot(formatAction('A1:C1', { fillColor: '#FF0000' }), ctx);
      expect(snapshot?.palette).toEqual([{ clearFill: true }, { fillColor: '#FFFF00' }]);
      expect(snapshot?.grid).toEqual([[0, 1, 0]]);
    });

    it('does not take a cell for unfilled just because it matches the corner, on a sheet filled edge to edge', async () => {
      const yellow = { fillPattern: 'Solid', fillColor: '#FFFF00' };
      const { ctx } = stubExcel({ fillPattern: 'null', corner: yellow, cell: () => yellow });
      const snapshot = await captureFormatSnapshot(formatAction('A1:B1', { fillColor: '#FF0000' }), ctx);
      expect(snapshot?.palette).toEqual([{ fillColor: '#FFFF00' }]);
    });

    it('refuses a white cell on such a sheet, where nothing says whether it is a fill', async () => {
      const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
      const { ctx } = stubExcel({
        fillPattern: 'null',
        corner: { fillPattern: 'Solid', fillColor: '#FFFF00' },
        cell: () => ({ fillPattern: 'Solid', fillColor: '#FFFFFF' }),
      });

      const snapshot = await captureFormatSnapshot(formatAction('A1:B1', { fillColor: '#FF0000' }), ctx);

      expect(snapshot).toMatchObject({ restorable: false });
      expect(snapshot?.reason).toMatch(/white fill could not be told from no fill/);
      // What Excel returned is logged, formatting only, so the next fix is not a guess.
      const logged = warn.mock.calls.find((call) => String(call[0]).includes('Format snapshot not taken'));
      expect(String(logged?.[2])).toContain('"unfilledReference":{"color":"#FFFF00"');
      warn.mockRestore();
    });
  });

  it('where Excel does report the pattern, the pattern decides, whatever the colour', async () => {
    // White with pattern 'None' is no fill; white with pattern 'Solid' is a white fill.
    const { ctx } = stubExcel({ cell: (_r, c) => (c === 1 ? { fillPattern: 'Solid', fillColor: '#FFFFFF' } : {}) });
    const snapshot = await captureFormatSnapshot(formatAction('A1:B1', { fillColor: '#FF0000' }), ctx);
    expect(snapshot?.palette).toEqual([{ clearFill: true }, { fillColor: '#FFFFFF' }]);
  });

  it('refuses borders without reading anything', async () => {
    const { ctx, worksheet } = stubExcel({ cell: () => ({}) });
    const snapshot = await captureFormatSnapshot(formatAction('A1:B2', { borders: { style: 'thin', edges: ['all'] } }), ctx);
    expect(snapshot).toMatchObject({ restorable: false });
    expect(worksheet.getRange).not.toHaveBeenCalled();
  });

  it('refuses a range with more cells than a snapshot may hold', async () => {
    const { ctx } = stubExcel({ cell: () => ({}) });
    const rangeWith = (cells: number) => {
      const worksheetRange = { rowIndex: 0, columnIndex: 0, rowCount: cells, columnCount: 1, load: vi.fn() };
      (ctx as unknown as { workbook: { worksheets: { getItem: () => { getRange: () => unknown } } } }).workbook.worksheets.getItem = () => ({
        getRange: () => worksheetRange,
      });
    };
    rangeWith(MAX_SNAPSHOT_CELLS + 1);
    const snapshot = await captureFormatSnapshot(formatAction('A:A', { bold: true }), ctx);
    expect(snapshot).toMatchObject({ restorable: false });
    expect(snapshot?.reason).toMatch(/more than/);
  });

  it('counts every range against one budget for the whole Accept, and refuses the one that goes over', async () => {
    const { ctx } = stubExcel({ cell: () => ({}) });
    const budget = { cells: 10 };

    const first = await captureFormatSnapshot(formatAction('A1:D2', { bold: true }), ctx, budget); // 8 cells
    const second = await captureFormatSnapshot(formatAction('A3:D3', { bold: true }), ctx, budget); // 4 more: over

    expect(first).toMatchObject({ restorable: true });
    expect(budget.cells).toBe(2);
    expect(second).toMatchObject({ restorable: false });
    expect(second?.reason).toMatch(/in all/);
  });

  it('refuses a range with more distinct formats than it is willing to send', async () => {
    // Every cell a different number format: the snapshot would be one format per cell.
    const { ctx } = stubExcel({ cell: (r, c) => ({ numberFormat: `0.${'0'.repeat(r)}"${c}"` }) });
    const wide = await captureFormatSnapshot(formatAction('A1:Z100', { numberFormat: '0' }), ctx);
    expect(wide).toMatchObject({ restorable: false });
    expect(wide?.reason).toMatch(/different formats/);
  });

  it('records an action that changes nothing it can read as a snapshot with nothing to restore', async () => {
    const { ctx } = stubExcel({ cell: () => ({}) });
    const snapshot = await captureFormatSnapshot(formatAction('A1:B2', { underline: true }), ctx);
    expect(snapshot).toMatchObject({ restorable: true, rowCount: 0, colCount: 0, palette: [], grid: [] });
  });

  it('turns an Excel failure into a refusal instead of throwing into the apply', async () => {
    const { ctx } = stubExcel({ cell: () => ({}) }, { sync: async () => { throw new Error('ItemNotFound'); } });
    const snapshot = await captureFormatSnapshot(formatAction('A1', { bold: true }), ctx);
    expect(snapshot).toMatchObject({ kind: 'format', restorable: false });
    expect(snapshot?.reason).toMatch(/ItemNotFound/);
  });

  it('returns null for an action that is not formatting', async () => {
    const { ctx } = stubExcel({ cell: () => ({}) });
    expect(await captureFormatSnapshot({ type: 'SET_CELL', sheetName: 'S', address: 'A1', value: 1 } as RichAction, ctx)).toBeNull();
  });
});

describe('captureFormatSnapshot — AUTOFIT_COLUMNS', () => {
  const autofit = (columns?: string[]) => ({ type: 'AUTOFIT_COLUMNS', sheetName: 'Monthly Summary', columns }) as unknown as RichAction;

  it('reads the width of each used column when none are named', async () => {
    const { ctx } = stubExcel({ cell: () => ({}), usedColumns: { columnIndex: 0, columnCount: 3 }, columnWidths: { 0: 64, 1: 90, 2: 64 } });
    const snapshot = await captureFormatSnapshot(autofit(), ctx);
    expect(snapshot).toEqual({
      kind: 'columns',
      sheetName: 'Monthly Summary',
      restorable: true,
      widths: [{ col: 0, width: 64 }, { col: 1, width: 90 }, { col: 2, width: 64 }],
    });
  });

  it('reads only the columns named', async () => {
    const { ctx } = stubExcel({ cell: () => ({}), columnWidths: { 1: 90, 3: 120 } });
    const snapshot = await captureFormatSnapshot(autofit(['B', 'D']), ctx);
    expect(snapshot?.widths).toEqual([{ col: 1, width: 90 }, { col: 3, width: 120 }]);
  });

  it('has nothing to restore on an empty sheet', async () => {
    const { ctx } = stubExcel({ cell: () => ({}), usedColumns: null });
    expect(await captureFormatSnapshot(autofit(), ctx)).toMatchObject({ restorable: true, widths: [] });
  });

  it('refuses a column name it does not recognise, and too many columns', async () => {
    const { ctx } = stubExcel({ cell: () => ({}), usedColumns: { columnIndex: 0, columnCount: MAX_SNAPSHOT_COLUMNS + 1 } });
    expect(await captureFormatSnapshot(autofit(['A1:B2']), ctx)).toMatchObject({ restorable: false });
    expect(await captureFormatSnapshot(autofit(), ctx)).toMatchObject({ restorable: false });
  });
});

describe('RichActionEngine.applyActions — formatSnapshots (TASKS.md #400)', () => {
  it('reports one snapshot per formatting action, taken before it ran, in the order applied', async () => {
    const { worksheet } = stubExcel({ cell: () => ({}), usedColumns: { columnIndex: 0, columnCount: 2 } });
    const order: string[] = [];
    worksheet.getRange.mockImplementation(((address: string) => {
      order.push(`getRange ${address}`);
      const b = bounds(address);
      const cell = { format: { font: { bold: false, italic: false, size: 11, name: 'x', color: '#000000' }, fill: { pattern: 'None', color: '#FFFFFF' }, horizontalAlignment: 'General', verticalAlignment: 'Bottom', wrapText: false } };
      return {
        ...b,
        load: vi.fn(),
        numberFormat: [['General']],
        getCellProperties: () => ({ value: [[cell]] }),
        format: {
          font: new Proxy({}, { set: (_t, key) => (order.push(`write ${String(key)}`), true) }),
          fill: { clear: vi.fn(), pattern: '', color: '' },
          borders: { getItem: vi.fn(() => ({})) },
        },
        select: vi.fn(),
      };
    }) as never);

    const engine = new RichActionEngine();
    const result = await engine.applyActions([
      formatAction('A3', { bold: true }),
      { type: 'AUTOFIT_COLUMNS', sheetName: 'Monthly Summary' } as unknown as RichAction,
    ]);

    expect(result.errors).toEqual([]);
    expect(result.formatSnapshots).toHaveLength(2);
    expect(result.formatSnapshots?.map((s) => s.kind)).toEqual(['format', 'columns']);
    expect(result.formatSnapshots?.every((s) => s.restorable)).toBe(true);
    // The read happened before the write: the first getRange (the read) precedes `write bold`.
    expect(order.indexOf('write bold')).toBeGreaterThan(order.indexOf('getRange A3'));
  });

  it('omits formatSnapshots when no formatting action ran', async () => {
    stubExcel({ cell: () => ({}) });
    const result = await new RichActionEngine().applyActions([
      { type: 'SET_CELL', sheetName: 'Monthly Summary', address: 'A1', value: 'x' } as unknown as RichAction,
    ]);
    expect(result.applied).toBe(1);
    expect(result.formatSnapshots).toBeUndefined();
  });

  it('does not report a snapshot for an action that failed', async () => {
    const { worksheet } = stubExcel({ cell: () => ({}) });
    const original = worksheet.getRange.getMockImplementation()!;
    let reads = 0;
    worksheet.getRange.mockImplementation(((address: string) => {
      reads += 1;
      // First call is the snapshot read; the second is the handler, which blows up.
      if (reads === 2) throw new Error('boom');
      return original(address);
    }) as never);

    const result = await new RichActionEngine().applyActions([formatAction('A1', { bold: true })]);
    expect(result.errors.join(' ')).toMatch(/boom/);
    expect(result.formatSnapshots).toBeUndefined();
  });
});
