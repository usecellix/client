import { CellValue, RichAction } from '@/action.types';
import { SheetAction } from '@/types/sheet-actions';
import { columnIndexToLetter } from './addressUtils';

function cellAddress(row: number, col: number): string {
  return `${columnIndexToLetter(col)}${row + 1}`;
}

function rangeAddress(
  row: number,
  col: number,
  rowCount: number,
  colCount: number,
): string {
  const endRow = row + rowCount - 1;
  const endCol = col + colCount - 1;
  const start = cellAddress(row, col);
  if (rowCount === 1 && colCount === 1) return start;
  return `${start}:${cellAddress(endRow, endCol)}`;
}

function sheetNameOf(action: SheetAction): string {
  return String(action.sheetName ?? '');
}

/** Convert legacy row/col SheetAction shapes to address-based RichAction. */
export function convertLegacyToRich(action: SheetAction): RichAction | RichAction[] | null {
  const sheetName = sheetNameOf(action);

  switch (action.type) {
    case 'SET_CELL':
      if (action.row === undefined || action.col === undefined) return null;
      return {
        type: 'SET_CELL',
        sheetName,
        address: cellAddress(action.row, action.col),
        value: action.value as CellValue,
        format: action.format as RichAction extends { type: 'SET_CELL' } ? never : never,
      } as RichAction;

    case 'CLEAR_CELL':
      if (action.row === undefined || action.col === undefined) return null;
      return {
        type: 'CLEAR_RANGE',
        sheetName,
        range: cellAddress(action.row, action.col),
        mode: 'contents',
      };

    case 'HIGHLIGHT_CELL':
      if (action.row === undefined || action.col === undefined) return null;
      return {
        type: 'HIGHLIGHT_CELL',
        sheetName,
        address: cellAddress(action.row, action.col),
        color: action.color,
      };

    case 'SET_FORMULA':
      if (action.row === undefined || action.col === undefined || action.formula === undefined) {
        return null;
      }
      return {
        type: 'SET_FORMULA',
        sheetName,
        address: cellAddress(action.row, action.col),
        formula: action.formula,
        format: action.format as RichAction extends { type: 'SET_FORMULA' } ? never : never,
      } as RichAction;

    case 'ADD_ROW':
      if (Array.isArray(action.data)) {
        return { type: 'APPEND_ROW', sheetName, values: action.data as CellValue[] };
      }
      if (typeof action.afterRow === 'number' && Array.isArray(action.values)) {
        return {
          type: 'ADD_ROW',
          sheetName,
          afterRow: action.afterRow,
          values: action.values as CellValue[],
          copyFormatFromRow: action.copyFormatFromRow as number | undefined,
        };
      }
      return null;

    case 'DELETE_ROW':
      if (action.row === undefined) return null;
      return { type: 'DELETE_ROW', sheetName, rows: [action.row + 1] };

    case 'INSERT_ROW':
      return {
        type: 'INSERT_ROW',
        sheetName,
        row: action.row ?? 0,
        count: action.count,
        position: action.position as 'above' | 'below' | undefined,
      };

    case 'INSERT_COLUMN':
      if (action.col === undefined) return null;
      return {
        type: 'INSERT_COLUMN',
        sheetName,
        beforeColumn: columnIndexToLetter(action.col),
        count: action.count ?? 1,
      };

    case 'DELETE_COLUMN':
      if (action.col === undefined) return null;
      return {
        type: 'DELETE_COLUMN',
        sheetName,
        columns: [columnIndexToLetter(action.col)],
      };

    case 'FORMAT_RANGE': {
      const row = action.row ?? 0;
      const col = action.col ?? 0;
      const rowCount = action.rowCount ?? 1;
      const colCount = action.colCount ?? 1;
      return {
        type: 'FORMAT_RANGE',
        sheetName,
        range: rangeAddress(row, col, rowCount, colCount),
        format: action.format as RichAction extends { type: 'FORMAT_RANGE' } ? never : never,
      } as RichAction;
    }

    case 'FILL_DOWN':
      if (action.row === undefined || action.col === undefined) return null;
      return {
        type: 'AUTO_FILL',
        sheetName,
        startAddress: cellAddress(action.row, action.col),
        direction: 'down',
        endRow: action.endRow !== undefined ? action.endRow + 1 : undefined,
      };

    case 'FILL_RIGHT':
      if (action.row === undefined || action.col === undefined) return null;
      return {
        type: 'AUTO_FILL',
        sheetName,
        startAddress: cellAddress(action.row, action.col),
        direction: 'right',
        endCol: action.endCol,
      };

    case 'WRITE_TABLE':
      return {
        type: 'WRITE_TABLE',
        sheetName,
        headers: (action.headers ?? []) as CellValue[],
        rows: (action.rows ?? []) as CellValue[][],
      };

    case 'CREATE_SHEET':
      return {
        type: 'ADD_SHEET',
        name: action.sheetName ?? action.name ?? 'New Sheet',
      };

    case 'RENAME_SHEET':
      return {
        type: 'RENAME_SHEET',
        oldName: action.sheetName ?? '',
        newName: action.newSheetName ?? action.sheetName ?? '',
      };

    case 'COPY_SHEET':
      return {
        type: 'COPY_SHEET',
        sourceName: action.sheetName ?? '',
        newName: action.newSheetName ?? `${action.sheetName ?? 'Sheet'} Copy`,
      };

    case 'HIDE_ROW':
    case 'UNHIDE_ROW':
    case 'SHOW_ROW':
      if (action.row === undefined) return null;
      return {
        type: action.type === 'HIDE_ROW' ? 'HIDE_ROW' : 'UNHIDE_ROW',
        sheetName,
        row: action.row,
        rowCount: action.rowCount ?? 1,
      } as unknown as RichAction;

    case 'HIDE_COLUMN':
    case 'UNHIDE_COLUMN':
    case 'SHOW_COLUMN':
      if (action.col === undefined) return null;
      return {
        type: action.type === 'HIDE_COLUMN' ? 'HIDE_COLUMN' : 'UNHIDE_COLUMN',
        sheetName,
        col: action.col,
        colCount: action.colCount ?? 1,
      } as unknown as RichAction;

    case 'SET_ROW_HEIGHT':
      if (action.row === undefined || action.height === undefined) return null;
      return {
        type: 'SET_ROW_HEIGHT',
        sheetName,
        row: action.row,
        height: action.height,
      } as unknown as RichAction;

    case 'SET_COLUMN_WIDTH':
      if (action.col === undefined || action.width === undefined) return null;
      return {
        type: 'SET_COLUMN_WIDTH',
        sheetName,
        col: action.col,
        width: action.width,
      } as unknown as RichAction;

    case 'FREEZE_PANES':
      return {
        type: 'FREEZE_PANES',
        sheetName,
        freezeRows: action.freezeRows ?? 1,
        freezeColumns: action.freezeColumns ?? 0,
      } as unknown as RichAction;

    case 'UNFREEZE_PANES':
      return { type: 'UNFREEZE_PANES', sheetName } as unknown as RichAction;

    case 'SET_ZOOM':
      if (typeof action.zoomPercent !== 'number') return null;
      return { type: 'SET_ZOOM', sheetName, zoomPercent: action.zoomPercent } as unknown as RichAction;

    case 'PROTECT_SHEET':
      return { type: 'PROTECT_SHEET', sheetName } as unknown as RichAction;

    case 'UNPROTECT_SHEET':
      return { type: 'UNPROTECT_SHEET', sheetName } as unknown as RichAction;

    case 'UNMERGE_CELLS': {
      const row = action.row ?? 0;
      const col = action.col ?? 0;
      return {
        type: 'UNMERGE_CELLS',
        sheetName,
        range: rangeAddress(row, col, action.rowCount ?? 1, action.colCount ?? 1),
      } as unknown as RichAction;
    }

    case 'HIDE_SHEET':
      return { type: 'HIDE_SHEET', sheetName } as unknown as RichAction;

    case 'SHOW_SHEET':
      return { type: 'SHOW_SHEET', sheetName } as unknown as RichAction;

    case 'SET_SHEET_COLOR':
      if (!action.color) return null;
      return { type: 'SET_SHEET_COLOR', sheetName, color: action.color } as unknown as RichAction;

    case 'ADD_COMMENT':
      if (action.row === undefined || action.col === undefined || !action.comment) return null;
      return {
        type: 'ADD_COMMENT',
        sheetName,
        address: cellAddress(action.row, action.col),
        comment: action.comment,
      } as unknown as RichAction;

    case 'DELETE_COMMENT':
      if (action.row === undefined || action.col === undefined) return null;
      return {
        type: 'DELETE_COMMENT',
        sheetName,
        address: cellAddress(action.row, action.col),
      } as unknown as RichAction;

    case 'MERGE_CELLS': {
      const row = action.row ?? 0;
      const col = action.col ?? 0;
      return {
        type: 'MERGE_CELLS',
        sheetName,
        range: rangeAddress(row, col, action.rowCount ?? 1, action.colCount ?? 1),
      };
    }

    case 'CLEAR_CONTENT':
    case 'CLEAR_FORMAT':
    case 'CLEAR_ALL': {
      const row = action.row ?? 0;
      const col = action.col ?? 0;
      const mode =
        action.type === 'CLEAR_CONTENT'
          ? 'contents'
          : action.type === 'CLEAR_FORMAT'
            ? 'formats'
            : 'all';
      return {
        type: 'CLEAR_RANGE',
        sheetName,
        range: rangeAddress(row, col, action.rowCount ?? 1, action.colCount ?? 1),
        mode,
      };
    }

    default:
      return null;
  }
}
