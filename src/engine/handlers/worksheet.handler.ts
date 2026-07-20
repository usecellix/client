import { RichAction } from '@/action.types';
import { resolveWorksheet } from '../sheetResolve';

/* global Excel */

type ActionRecord = Record<string, unknown> & { type: string };

function asActionRecord(action: RichAction): ActionRecord {
  return action as unknown as ActionRecord;
}

export async function handleWorksheetAction(
  action: RichAction,
  ctx: Excel.RequestContext,
): Promise<boolean> {
  const record = asActionRecord(action);
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
    case 'SET_ROW_HEIGHT': {
      const row = Number(record.row);
      const height = Number(record.height);
      if (!Number.isInteger(row) || row < 0 || !Number.isFinite(height)) return true;
      sheet.getRangeByIndexes(row, 0, 1, 1).format.rowHeight = height;
      await ctx.sync();
      return true;
    }
    case 'SET_COLUMN_WIDTH': {
      const col = Number(record.col);
      const width = Number(record.width);
      if (!Number.isInteger(col) || col < 0 || !Number.isFinite(width)) return true;
      sheet.getRangeByIndexes(0, col, 1, 1).format.columnWidth = width;
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
    case 'SET_ZOOM': {
      const zoom = Number(record.zoomPercent);
      if (!Number.isFinite(zoom)) return true;
      sheet.sheetView.zoom = Math.max(10, Math.min(400, Math.round(zoom)));
      await ctx.sync();
      return true;
    }
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
    case 'ADD_COMMENT':
      if (typeof record.address !== 'string' || typeof record.comment !== 'string') return true;
      sheet.getRange(record.address).getComment().add(record.comment);
      await ctx.sync();
      return true;
    case 'DELETE_COMMENT':
      if (typeof record.address !== 'string') return true;
      sheet.getRange(record.address).getComment().delete();
      await ctx.sync();
      return true;
    default:
      return false;
  }
}
