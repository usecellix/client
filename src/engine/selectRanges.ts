import { RichAction } from '@/action.types';
import { SheetAction } from '@/types/sheet-actions';
import {
  columnIndexToLetter,
  columnLetterToIndex,
  parseCellAddress,
  parseRangeAddress,
} from './addressUtils';

/* global Excel */

export type SelectableBounds = {
  sheetName: string;
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
};

/** Build an A1 range string from 0-based bounds. */
export function boundsToAddress(bounds: SelectableBounds): string {
  const start = `${columnIndexToLetter(bounds.col)}${bounds.row + 1}`;
  if (bounds.rowCount <= 1 && bounds.colCount <= 1) return start;
  const end = `${columnIndexToLetter(bounds.col + bounds.colCount - 1)}${bounds.row + bounds.rowCount}`;
  return `${start}:${end}`;
}

/**
 * Resolve the sheet range that an action will affect (best-effort, before or after apply).
 * Used to mouse-select edited cells/rows/columns — no fill colors.
 */
export function resolveActionSelectBounds(
  action: SheetAction | RichAction,
  fallbackSheet: string,
): SelectableBounds | null {
  const stamped = (action as { __appliedBounds?: SelectableBounds }).__appliedBounds;
  if (stamped) return stamped;

  const sheetName = String(
    (action as SheetAction).sheetName ??
      (action as { destSheet?: string }).destSheet ??
      (action as { sourceSheet?: string }).sourceSheet ??
      fallbackSheet,
  );
  if (!sheetName) return null;

  // NB: this intersects `type` down to the tags common to both unions, so any
  // rich-only action handled below must also appear in SheetActionType.
  const rich = action as RichAction & SheetAction;

  switch (rich.type) {
    case 'SET_CELL':
    case 'SET_FORMULA':
    case 'HIGHLIGHT_CELL': {
      if (typeof rich.address === 'string') {
        const cell = parseCellAddress(rich.address);
        if (cell) return { sheetName, row: cell.row, col: cell.col, rowCount: 1, colCount: 1 };
      }
      if (rich.row !== undefined && rich.col !== undefined) {
        return {
          sheetName,
          row: rich.row,
          col: rich.col,
          rowCount: rich.rowCount ?? 1,
          colCount: rich.colCount ?? 1,
        };
      }
      return null;
    }
    case 'FORMAT_RANGE':
    case 'CLEAR_RANGE':
    case 'MERGE_CELLS':
    case 'CONDITIONAL_FORMAT': {
      if (typeof rich.range === 'string') {
        const parsed = parseRangeAddress(rich.range);
        if (parsed) return { sheetName, ...parsed };
      }
      if (rich.row !== undefined && rich.col !== undefined) {
        return {
          sheetName,
          row: rich.row,
          col: rich.col,
          rowCount: rich.rowCount ?? 1,
          colCount: rich.colCount ?? 1,
        };
      }
      return null;
    }
    case 'INSERT_COLUMN': {
      if (typeof rich.beforeColumn === 'string') {
        return {
          sheetName,
          row: 0,
          col: columnLetterToIndex(rich.beforeColumn),
          rowCount: 1,
          colCount: rich.count ?? 1,
        };
      }
      if (rich.col !== undefined) {
        return {
          sheetName,
          row: 0,
          col: rich.col,
          rowCount: Math.max(rich.rowCount ?? 1, 1),
          colCount: rich.count ?? 1,
        };
      }
      return null;
    }
    case 'ADD_ROW':
    case 'APPEND_ROW':
    case 'INSERT_ROW': {
      if (rich.row !== undefined) {
        return { sheetName, row: rich.row, col: 0, rowCount: rich.count ?? 1, colCount: 1 };
      }
      if (typeof rich.afterRow === 'number') {
        return { sheetName, row: rich.afterRow, col: 0, rowCount: 1, colCount: 1 };
      }
      return null;
    }
    case 'DELETE_ROW': {
      if (Array.isArray((rich as { rows?: number[] }).rows) && (rich as { rows: number[] }).rows.length) {
        const rows = (rich as { rows: number[] }).rows;
        const min = Math.min(...rows) - 1;
        const max = Math.max(...rows) - 1;
        return { sheetName, row: min, col: 0, rowCount: max - min + 1, colCount: 1 };
      }
      if (rich.row !== undefined) {
        return { sheetName, row: rich.row, col: 0, rowCount: 1, colCount: 1 };
      }
      return null;
    }
    case 'WRITE_TABLE': {
      const headers = (rich as { headers?: unknown[] }).headers ?? [];
      const rows = (rich as { rows?: unknown[][] }).rows ?? [];
      return {
        sheetName,
        row: 0,
        col: 0,
        rowCount: Math.max(1 + rows.length, 1),
        colCount: Math.max(headers.length, 1),
      };
    }
    case 'COPY_FILTERED_RANGE':
    case 'MOVE_RANGE':
    case 'AGGREGATE_TABLE': {
      const destSheet = String(
        (rich as { destSheet?: string }).destSheet ?? sheetName,
      );
      const start = parseCellAddress(
        String((rich as { destStartCell?: string }).destStartCell ?? 'A1'),
      );
      if (!start) return null;
      return { sheetName: destSheet, row: start.row, col: start.col, rowCount: 1, colCount: 1 };
    }
    case 'FORMAT_MATCHING_ROWS':
    case 'SET_MATCHING_ROWS': {
      if (typeof rich.range === 'string') {
        const parsed = parseRangeAddress(rich.range);
        if (parsed) return { sheetName, ...parsed };
      }
      return null;
    }
    case 'BATCH_SET': {
      const ops = (rich as { operations?: { address: string }[] }).operations ?? [];
      const cells = ops
        .map((op) => parseCellAddress(op.address))
        .filter((c): c is { row: number; col: number } => Boolean(c));
      if (!cells.length) return null;
      const minR = Math.min(...cells.map((c) => c.row));
      const maxR = Math.max(...cells.map((c) => c.row));
      const minC = Math.min(...cells.map((c) => c.col));
      const maxC = Math.max(...cells.map((c) => c.col));
      return {
        sheetName,
        row: minR,
        col: minC,
        rowCount: maxR - minR + 1,
        colCount: maxC - minC + 1,
      };
    }
    case 'SORT_RANGE': {
      if (typeof rich.range === 'string') {
        const parsed = parseRangeAddress(rich.range);
        if (parsed) return { sheetName, ...parsed };
      }
      return null;
    }
    default:
      return null;
  }
}

/** Select (mouse-style) the ranges touched by actions — no fill/highlight colors. */
export async function selectActionRanges(
  actions: Array<SheetAction | RichAction>,
  ctx: Excel.RequestContext,
): Promise<void> {
  if (!actions.length) return;

  const active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load('name');
  await ctx.sync();

  const boundsList: SelectableBounds[] = [];
  for (const action of actions) {
    const bounds = resolveActionSelectBounds(action, active.name);
    if (bounds) boundsList.push(bounds);
  }
  if (!boundsList.length) return;

  // Select the first sheet's union span (Excel single selection).
  const primarySheet = boundsList[0].sheetName;
  const sameSheet = boundsList.filter((b) => b.sheetName === primarySheet);
  const minR = Math.min(...sameSheet.map((b) => b.row));
  const minC = Math.min(...sameSheet.map((b) => b.col));
  const maxR = Math.max(...sameSheet.map((b) => b.row + b.rowCount - 1));
  const maxC = Math.max(...sameSheet.map((b) => b.col + b.colCount - 1));

  const sheet =
    primarySheet === active.name
      ? active
      : ctx.workbook.worksheets.getItem(primarySheet);
  if (primarySheet !== active.name) {
    sheet.activate();
  }

  const range = sheet.getRangeByIndexes(
    minR,
    minC,
    maxR - minR + 1,
    maxC - minC + 1,
  );
  range.select();
  await ctx.sync();
}

/** Attach bounds onto an action object for later selection (mutates). */
export function stampAppliedBounds(
  action: RichAction,
  bounds: SelectableBounds,
): void {
  (action as RichAction & { __appliedBounds?: SelectableBounds }).__appliedBounds = bounds;
}
