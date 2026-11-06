import type { RichAction } from '@/action.types';
import { columnIndexToLetter, stripSheetPrefix, isLocalRangeAddress } from './addressUtils';
import {
  captureFormatSnapshot,
  MAX_SNAPSHOT_TOTAL_CELLS,
  type FormatSnapshot,
  type SnapshotBudget,
  type SnapshotCellFormat,
} from './formatSnapshot';
import { extendRangeToUsedRows } from './handlers/range.handler';
import { applyFormat } from './handlers/format.handler';
import { findMatchingRowOffsets, resolveFilterColumnIndex } from './rangeFilter';
import { resolveWorksheet } from './sheetResolve';

/* global Excel */

/**
 * P1.6 / B3 — session-only snapshot for the row-level actions Excel cannot undo
 * by itself and the server cannot rebuild (DELETE_MATCHING_ROWS, SET_MATCHING_ROWS).
 *
 * What it holds: the exact cells (formulas/constants, number formats, and the
 * row formatting) that the action is about to remove or overwrite, read straight
 * off the workbook just before the action runs.
 *
 * Where it lives: in this module's memory ONLY (`sessionUndoStore`). It is never
 * sent to the server, never written to the workbook, never logged — client data
 * retention rule (§1.5). Consequence: undo works until the task pane is reloaded;
 * after that the change cannot be undone from CELLIX (Excel's own Ctrl+Z still can).
 */

/** Hard limits so a snapshot can never become the heavy part of an Accept. */
export const MAX_SNAPSHOT_ROWS = 2_000;
export const MAX_SNAPSHOT_ROW_CELLS = 60_000;

/** Action types whose undo is served by this module. Mirrored in Server reversibility-catalog.ts (parity spec). */
export const SESSION_UNDOABLE_ACTION_TYPES = ['DELETE_MATCHING_ROWS', 'SET_MATCHING_ROWS'] as const;

export function withoutSessionUndoable(types: string[] | undefined): string[] {
  const list = types ?? [];
  return list.filter((t) => !(SESSION_UNDOABLE_ACTION_TYPES as readonly string[]).includes(t));
}

export interface RowSnapshotRow {
  /** 0-based worksheet row index at capture time (before the action ran). */
  index: number;
  /** Formulas/constants exactly as Excel returns them (a constant stays a constant). */
  formulas: unknown[];
  numberFormats: string[];
}

export interface RowSnapshot {
  kind: 'deleted-rows' | 'set-cells';
  sheetName: string;
  restorable: boolean;
  reason?: string;
  /** First column of every row's content (0 for whole rows, the target column for set-cells). */
  colStart: number;
  colCount: number;
  rows: RowSnapshotRow[];
  /** Row formatting (deleted-rows only). A refused one is skipped: values still come back. */
  formats: FormatSnapshot[];
}

type DeleteAction = Extract<RichAction, { type: 'DELETE_MATCHING_ROWS' }>;
type SetAction = Extract<RichAction, { type: 'SET_MATCHING_ROWS' }>;

const FORMAT_PROPS = {
  bold: true,
  italic: true,
  fontSize: 11,
  fontName: 'x',
  fontColor: '#000000',
  fillColor: '#000000',
  numberFormat: '@',
  horizontalAlignment: 'left',
  verticalAlignment: 'top',
  wrapText: false,
} as const;

function refusal(kind: RowSnapshot['kind'], sheetName: string, reason: string): RowSnapshot {
  return { kind, sheetName, restorable: false, reason, colStart: 0, colCount: 0, rows: [], formats: [] };
}

/** Contiguous runs of an ascending row-index list: [3,4,5,9] → [[3,5],[9,9]]. */
export function contiguousRuns(indices: number[]): Array<[number, number]> {
  const sorted = [...indices].sort((a, b) => a - b);
  const runs: Array<[number, number]> = [];
  for (const i of sorted) {
    const last = runs[runs.length - 1];
    if (last && i === last[1] + 1) last[1] = i;
    else runs.push([i, i]);
  }
  return runs;
}

/**
 * The snapshot for a row-level action, or null for any other action. Never
 * throws: a failure becomes a snapshot that refuses the undo with a reason.
 */
export async function captureRowSnapshot(
  action: RichAction,
  ctx: Excel.RequestContext,
): Promise<RowSnapshot | null> {
  if (action.type !== 'DELETE_MATCHING_ROWS' && action.type !== 'SET_MATCHING_ROWS') return null;
  const kind: RowSnapshot['kind'] = action.type === 'DELETE_MATCHING_ROWS' ? 'deleted-rows' : 'set-cells';
  try {
    return action.type === 'DELETE_MATCHING_ROWS'
      ? await captureDeletedRows(action, ctx)
      : await captureSetCells(action, ctx);
  } catch (error) {
    return refusal(
      kind,
      action.sheetName,
      `previous contents could not be read: ${error instanceof Error ? error.message : String(error)}`.slice(0, 200),
    );
  }
}

async function scanMatches(
  action: DeleteAction | SetAction,
  ctx: Excel.RequestContext,
): Promise<{ sheet: Excel.Worksheet; rowIndex: number; columnIndex: number; rows: unknown[][]; offsets: number[]; hasHeaders: boolean }> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const local = stripSheetPrefix(action.range);
  if (!isLocalRangeAddress(local)) throw new Error(`Invalid source range "${action.range}"`);
  // The same widening the handler applies, so both see the same rows.
  const range = await extendRangeToUsedRows(sheet, local, ctx);
  range.load(['values', 'rowIndex', 'columnIndex', 'columnCount']);
  await ctx.sync();
  const rows = (range.values ?? []) as unknown[][];
  const hasHeaders = action.hasHeaders !== false;
  let offsets: number[] = [];
  if (rows.length > 0) {
    if (action.type === 'DELETE_MATCHING_ROWS') {
      const firstDataOffset = hasHeaders ? 1 : 0;
      offsets = action.filter
        ? findMatchingRowOffsets(rows, hasHeaders, action.filter)
        : rows.reduce<number[]>((blanks, row, offset) => {
            if (offset < firstDataOffset) return blanks;
            if (row.every((cell) => cell === null || String(cell ?? '').trim() === '')) blanks.push(offset);
            return blanks;
          }, []);
    } else {
      offsets = action.filter
        ? findMatchingRowOffsets(rows, hasHeaders, action.filter)
        : Array.from({ length: Math.max(rows.length - (hasHeaders ? 1 : 0), 0) }, (_, i) => (hasHeaders ? i + 1 : i));
    }
  }
  return { sheet, rowIndex: range.rowIndex, columnIndex: range.columnIndex, rows, offsets, hasHeaders };
}

async function captureDeletedRows(action: DeleteAction, ctx: Excel.RequestContext): Promise<RowSnapshot> {
  const scan = await scanMatches(action, ctx);
  if (scan.offsets.length === 0) {
    return { kind: 'deleted-rows', sheetName: action.sheetName, restorable: true, colStart: 0, colCount: 0, rows: [], formats: [] };
  }
  if (scan.offsets.length > MAX_SNAPSHOT_ROWS) {
    return refusal('deleted-rows', action.sheetName, `more than ${MAX_SNAPSHOT_ROWS} rows match`);
  }

  // Whole rows go, so every used column is captured, from column A.
  const used = scan.sheet.getUsedRange();
  used.load(['columnIndex', 'columnCount']);
  await ctx.sync();
  const colCount = (used.columnIndex ?? 0) + (used.columnCount ?? 0);
  if (scan.offsets.length * colCount > MAX_SNAPSHOT_ROW_CELLS) {
    return refusal('deleted-rows', action.sheetName, `more than ${MAX_SNAPSHOT_ROW_CELLS} cells would be removed`);
  }

  const indices = scan.offsets.map((o) => scan.rowIndex + o).sort((a, b) => a - b);
  const ranges = indices.map((rowIdx) => {
    const r = scan.sheet.getRangeByIndexes(rowIdx, 0, 1, colCount);
    r.load(['formulas', 'numberFormat']);
    return r;
  });
  await ctx.sync();
  const rows: RowSnapshotRow[] = indices.map((index, i) => ({
    index,
    formulas: (ranges[i].formulas?.[0] ?? []) as unknown[],
    numberFormats: ((ranges[i].numberFormat?.[0] ?? []) as unknown[]).map((f) => String(f ?? 'General')),
  }));

  // Row formatting, one read per contiguous run. A run that cannot be read exactly
  // is recorded as refused and skipped at restore — values still come back.
  const budget: SnapshotBudget = { cells: MAX_SNAPSHOT_TOTAL_CELLS };
  const formats: FormatSnapshot[] = [];
  for (const [from, to] of contiguousRuns(indices)) {
    const address = `A${from + 1}:${columnIndexToLetter(colCount - 1)}${to + 1}`;
    const snap = await captureFormatSnapshot(
      { type: 'FORMAT_RANGE', sheetName: action.sheetName, range: address, format: { ...FORMAT_PROPS } } as unknown as RichAction,
      ctx,
      budget,
    );
    if (snap) formats.push(snap);
  }

  return { kind: 'deleted-rows', sheetName: action.sheetName, restorable: true, colStart: 0, colCount, rows, formats };
}

async function captureSetCells(action: SetAction, ctx: Excel.RequestContext): Promise<RowSnapshot> {
  const scan = await scanMatches(action, ctx);
  if (scan.offsets.length === 0) {
    return { kind: 'set-cells', sheetName: action.sheetName, restorable: true, colStart: 0, colCount: 1, rows: [], formats: [] };
  }
  if (scan.offsets.length > MAX_SNAPSHOT_ROWS) {
    return refusal('set-cells', action.sheetName, `more than ${MAX_SNAPSHOT_ROWS} rows match`);
  }
  const headerRow = scan.hasHeaders ? scan.rows[0] : null;
  if (!headerRow) return refusal('set-cells', action.sheetName, 'no header row to find the target column');
  const targetCol = scan.columnIndex + resolveFilterColumnIndex(headerRow, action.targetColumn);

  const indices = scan.offsets.map((o) => scan.rowIndex + o);
  const cells = indices.map((rowIdx) => {
    const c = scan.sheet.getRangeByIndexes(rowIdx, targetCol, 1, 1);
    c.load(['formulas', 'numberFormat']);
    return c;
  });
  await ctx.sync();
  const rows: RowSnapshotRow[] = indices.map((index, i) => ({
    index,
    formulas: [cells[i].formulas?.[0]?.[0] ?? ''],
    numberFormats: [String(cells[i].numberFormat?.[0]?.[0] ?? 'General')],
  }));
  return { kind: 'set-cells', sheetName: action.sheetName, restorable: true, colStart: targetCol, colCount: 1, rows, formats: [] };
}

// ---- Restore -----------------------------------------------------------------

/** Splits a grid of palette indices into rectangles (same algorithm as Server/src/audit/format-snapshot.ts). */
export function gridToRects(grid: number[][]): Array<{ row: number; col: number; rowCount: number; colCount: number; index: number }> {
  type Rect = { row: number; col: number; rowCount: number; colCount: number; index: number };
  const rects: Rect[] = [];
  let open = new Map<string, Rect>();
  grid.forEach((line, row) => {
    const next = new Map<string, Rect>();
    let start = 0;
    for (let col = 1; col <= line.length; col += 1) {
      if (col < line.length && line[col] === line[start]) continue;
      const key = `${start}:${col - 1}:${line[start]}`;
      const carried = open.get(key);
      if (carried) {
        carried.rowCount += 1;
        next.set(key, carried);
      } else {
        const rect = { row, col: start, rowCount: 1, colCount: col - start, index: line[start] };
        rects.push(rect);
        next.set(key, rect);
      }
      start = col;
    }
    open = next;
  });
  return rects;
}

export interface RestoreReport {
  rowsRestored: number;
  cellsRestored: number;
  /** Row-format runs that could not be read exactly at capture time (values were still restored). */
  formatRunsSkipped: number;
}

/** Puts back what one snapshot recorded. Throws if the snapshot cannot be restored. */
export async function restoreRowSnapshot(snapshot: RowSnapshot, ctx: Excel.RequestContext): Promise<RestoreReport> {
  if (!snapshot.restorable) {
    throw new Error(snapshot.reason ? `Cannot undo: ${snapshot.reason}` : 'Cannot undo this change.');
  }
  const sheet = resolveWorksheet(ctx, snapshot.sheetName);
  const rows = [...snapshot.rows].sort((a, b) => a.index - b.index);
  if (rows.length === 0) return { rowsRestored: 0, cellsRestored: 0, formatRunsSkipped: 0 };

  if (snapshot.kind === 'deleted-rows') {
    // Ascending: each insert lands at its ORIGINAL index because every earlier row is already back.
    for (const row of rows) {
      sheet.getRangeByIndexes(row.index, 0, 1, 1).getEntireRow().insert(Excel.InsertShiftDirection.down);
    }
    await ctx.sync();
  }

  for (const row of rows) {
    const target = sheet.getRangeByIndexes(row.index, snapshot.colStart, 1, snapshot.colCount);
    target.numberFormat = [row.numberFormats];
    target.formulas = [row.formulas as (string | number | boolean)[]];
  }
  await ctx.sync();

  let formatRunsSkipped = 0;
  for (const fmt of snapshot.formats) {
    if (!fmt.restorable) {
      formatRunsSkipped += 1;
      continue;
    }
    for (const rect of gridToRects(fmt.grid ?? [])) {
      const format = fmt.palette?.[rect.index] as SnapshotCellFormat | undefined;
      if (!format || Object.keys(format).length === 0) continue;
      const range = sheet.getRangeByIndexes(
        (fmt.row ?? 0) + rect.row,
        (fmt.col ?? 0) + rect.col,
        rect.rowCount,
        rect.colCount,
      );
      applyFormat(range, format as Parameters<typeof applyFormat>[1]);
    }
  }
  await ctx.sync();

  return {
    rowsRestored: rows.length,
    cellsRestored: rows.length * snapshot.colCount,
    formatRunsSkipped,
  };
}

// ---- Session store -------------------------------------------------------------

/**
 * changeSetId → snapshots taken when that change set was applied. Module memory
 * only (see header). Cleared by a task-pane reload by design.
 */
const store = new Map<string, RowSnapshot[]>();

export const sessionUndoStore = {
  register(changeSetId: string, snapshots: RowSnapshot[] | undefined): void {
    if (!changeSetId || !snapshots || snapshots.length === 0) return;
    store.set(changeSetId, snapshots);
  },
  has(changeSetId: string): boolean {
    return store.has(changeSetId);
  },
  /** True when every snapshot recorded for the change set can be restored. */
  isRestorable(changeSetId: string): boolean {
    const snaps = store.get(changeSetId);
    return !!snaps && snaps.every((s) => s.restorable);
  },
  clear(changeSetId?: string): void {
    if (changeSetId) store.delete(changeSetId);
    else store.clear();
  },
  /** Restores newest-first, so two actions that touched the same rows unwind in order. */
  async restore(changeSetId: string): Promise<RestoreReport> {
    const snaps = store.get(changeSetId);
    if (!snaps) throw new Error('No in-session undo data for this change (the task pane was reloaded).');
    const total: RestoreReport = { rowsRestored: 0, cellsRestored: 0, formatRunsSkipped: 0 };
    await Excel.run(async (ctx) => {
      for (const snap of [...snaps].reverse()) {
        const report = await restoreRowSnapshot(snap, ctx);
        total.rowsRestored += report.rowsRestored;
        total.cellsRestored += report.cellsRestored;
        total.formatRunsSkipped += report.formatRunsSkipped;
      }
    });
    store.delete(changeSetId);
    return total;
  },
};
