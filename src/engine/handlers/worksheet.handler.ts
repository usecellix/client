import { RichAction } from '@/action.types';
import { resolveWorksheet } from '../sheetResolve';
import { resolveFilterColumnIndex } from '@/shared/rangeFilter';

/* global Excel */

type ActionRecord = Record<string, unknown> & { type: string };

function asActionRecord(action: RichAction): ActionRecord {
  return action as unknown as ActionRecord;
}

/**
 * Exactly the action types the switch below handles.
 *
 * This gate is load-bearing, not a micro-optimisation. The engine tries this
 * handler FIRST for every action, and resolveWorksheet() calls
 * `worksheets.getItem(name)`, which queues a lookup on the Office.js request
 * context. Resolving before knowing whether we handle the type meant an
 * ADD_SHEET/CREATE_SHEET for a sheet that does not exist yet (the normal case —
 * it is about to be created) queued a lookup for a missing sheet. The switch
 * then fell through to `default`, the sheet was created correctly by a later
 * handler, but the poisoned lookup stayed in the batch and the next ctx.sync()
 * failed the whole run with "The requested resource doesn't exist."
 */
const WORKSHEET_ACTION_TYPES = new Set<string>([
  'HIDE_ROW',
  'UNHIDE_ROW',
  'HIDE_COLUMN',
  'UNHIDE_COLUMN',
  'SET_ROW_HEIGHT',
  'SET_COLUMN_WIDTH',
  'FREEZE_PANES',
  'UNFREEZE_PANES',
  'AUTO_FILTER',
  'SET_ZOOM',
  'PROTECT_SHEET',
  'UNPROTECT_SHEET',
  'UNMERGE_CELLS',
  'HIDE_SHEET',
  'SHOW_SHEET',
  'SET_SHEET_COLOR',
  'ADD_COMMENT',
  'DELETE_COMMENT',
]);

export async function handleWorksheetAction(
  action: RichAction,
  ctx: Excel.RequestContext,
): Promise<boolean> {
  const record = asActionRecord(action);
  // Resolve only after confirming we handle this type — see WORKSHEET_ACTION_TYPES.
  if (!WORKSHEET_ACTION_TYPES.has(record.type)) return false;
  const sheet = resolveWorksheet(ctx, String(record.sheetName ?? ''));

  switch (record.type) {
    case 'HIDE_ROW':
    case 'UNHIDE_ROW': {
      const row = Number(record.row);
      if (!Number.isInteger(row) || row < 0) return true;
      const rowCount = Number(record.rowCount ?? 1);
      sheet.getRangeByIndexes(row, 0, Math.max(1, rowCount), 1).rowHidden = record.type === 'HIDE_ROW';
      await ctx.sync();
      return true;
    }
    case 'HIDE_COLUMN':
    case 'UNHIDE_COLUMN': {
      const col = Number(record.col);
      if (!Number.isInteger(col) || col < 0) return true;
      const colCount = Number(record.colCount ?? 1);
      sheet.getRangeByIndexes(0, col, 1, Math.max(1, colCount)).columnHidden =
        record.type === 'HIDE_COLUMN';
      await ctx.sync();
      return true;
    }
    // rowCount/colCount were ignored, so "set the height of rows 2 to 5" only
    // ever resized row 2 — the shortcut router has always sent the count.
    // TASKS.md #216.
    case 'SET_ROW_HEIGHT': {
      const row = Number(record.row);
      const height = Number(record.height);
      if (!Number.isInteger(row) || row < 0 || !Number.isFinite(height)) return true;
      const rowCount = Math.max(1, Number(record.rowCount ?? 1));
      sheet.getRangeByIndexes(row, 0, rowCount, 1).format.rowHeight = height;
      await ctx.sync();
      return true;
    }
    case 'SET_COLUMN_WIDTH': {
      const col = Number(record.col);
      const width = Number(record.width);
      if (!Number.isInteger(col) || col < 0 || !Number.isFinite(width)) return true;
      const colCount = Math.max(1, Number(record.colCount ?? 1));
      sheet.getRangeByIndexes(0, col, 1, colCount).format.columnWidth = width;
      await ctx.sync();
      return true;
    }
    case 'FREEZE_PANES': {
      const freezeRows = Math.max(0, Number(record.freezeRows ?? 1));
      const freezeColumns = Math.max(0, Number(record.freezeColumns ?? 0));
      if (freezeRows > 0) sheet.freezePanes.freezeRows(freezeRows);
      if (freezeColumns > 0) sheet.freezePanes.freezeColumns(freezeColumns);
      await ctx.sync();
      return true;
    }
    case 'UNFREEZE_PANES':
      sheet.freezePanes.unfreeze();
      await ctx.sync();
      return true;
    case 'AUTO_FILTER': {
      if (typeof record.range !== 'string' || !record.range) return true;
      const filter = record.filter as
        | { column?: string; operator?: string; value?: string | number }
        | undefined;

      if (!filter?.column || !filter.operator) {
        // No condition — just the dropdown arrows (guide T2.1 "Add AutoFilter").
        sheet.autoFilter.apply(record.range);
        await ctx.sync();
        return true;
      }

      // "Show only rows where X" (guide T2.2) needs a real criterion, not only
      // the arrows — otherwise every row stays visible. TASKS.md #221.
      const range = sheet.getRange(record.range);
      range.load('values');
      await ctx.sync();

      const rows = (range.values ?? []) as unknown[][];
      const hasHeaders = record.hasHeaders !== false;
      const headerRow = hasHeaders ? rows[0] : null;
      if (!headerRow) {
        sheet.autoFilter.apply(record.range);
        await ctx.sync();
        return true;
      }

      let columnIndex: number;
      try {
        columnIndex = resolveFilterColumnIndex(headerRow, filter.column);
      } catch {
        sheet.autoFilter.apply(record.range);
        await ctx.sync();
        return true;
      }

      const criteria = buildAutoFilterCriteria(filter.operator, filter.value);
      if (!criteria) {
        // Length/regex conditions have no Excel AutoFilter equivalent — dropdowns only.
        sheet.autoFilter.apply(record.range);
        await ctx.sync();
        return true;
      }

      sheet.autoFilter.apply(record.range, columnIndex, criteria);
      await ctx.sync();
      return true;
    }
    case 'SET_ZOOM':
      // The Excel JavaScript API exposes no view-zoom control (Worksheet has only
      // namedSheetViews, and pageLayout.zoom applies to printing). This previously
      // assigned to a non-existent sheetView and failed with an opaque TypeError.
      throw new Error(
        'Zoom cannot be changed from an add-in — the Excel JavaScript API does not expose worksheet view zoom.',
      );
    case 'PROTECT_SHEET':
      sheet.protection.protect();
      await ctx.sync();
      return true;
    case 'UNPROTECT_SHEET':
      sheet.protection.unprotect();
      await ctx.sync();
      return true;
    case 'UNMERGE_CELLS':
      if (typeof record.range !== 'string' || !record.range) return true;
      sheet.getRange(record.range).unmerge();
      await ctx.sync();
      return true;
    case 'HIDE_SHEET':
      sheet.visibility = Excel.SheetVisibility.hidden;
      await ctx.sync();
      return true;
    case 'SHOW_SHEET':
      sheet.visibility = Excel.SheetVisibility.visible;
      await ctx.sync();
      return true;
    case 'SET_SHEET_COLOR':
      if (typeof record.color !== 'string' || !record.color) return true;
      sheet.tabColor = record.color;
      await ctx.sync();
      return true;
    // Comments hang off the workbook, not the range — Range has no getComment().
    case 'ADD_COMMENT':
      if (typeof record.address !== 'string' || typeof record.comment !== 'string') return true;
      ctx.workbook.comments.add(sheet.getRange(record.address), record.comment);
      await ctx.sync();
      return true;
    case 'DELETE_COMMENT': {
      if (typeof record.address !== 'string') return true;
      const comment = ctx.workbook.comments.getItemByCell(sheet.getRange(record.address));
      comment.delete();
      await ctx.sync();
      return true;
    }
    default:
      return false;
  }
}

/**
 * Translate the shared RangeFilterSpec DSL (already used by
 * SET_MATCHING_ROWS/DELETE_MATCHING_ROWS/FORMAT_MATCHING_ROWS) into a real
 * Excel.FilterCriteria — the same condition language the model already knows,
 * reused instead of inventing a second one for AUTO_FILTER. `equals` uses
 * FilterOn.values (an exact discrete-value match, what "supplier is ABC
 * Traders" means); everything else is a custom criterion string, Excel's own
 * filter syntax. lengthEquals/lengthNotEquals/matchesRegex/notMatchesRegex
 * have no AutoFilter equivalent and return null so the caller falls back to
 * plain dropdowns rather than silently applying the wrong condition.
 * TASKS.md #221.
 */
function buildAutoFilterCriteria(
  operator: string,
  value: string | number | undefined,
): Excel.FilterCriteria | null {
  if (value === undefined) return null;

  switch (operator) {
    case 'equals':
      return { filterOn: Excel.FilterOn.values, values: [String(value)] };
    case 'notEquals':
      return { filterOn: Excel.FilterOn.custom, criterion1: `<>${value}` };
    case 'contains':
      return { filterOn: Excel.FilterOn.custom, criterion1: `*${value}*` };
    case 'greaterThan':
      return { filterOn: Excel.FilterOn.custom, criterion1: `>${value}` };
    case 'lessThan':
      return { filterOn: Excel.FilterOn.custom, criterion1: `<${value}` };
    default:
      return null;
  }
}
