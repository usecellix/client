export interface FormatSpec {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  horizontalAlignment?: 'left' | 'center' | 'right';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numberFormat?: string;
  borders?: 'all' | 'outer' | 'bottom' | 'none';
}

export type SheetActionType =
  | 'SET_CELL'
  | 'CLEAR_CELL'
  | 'HIGHLIGHT_CELL'
  | 'SET_FORMULA'
  | 'ADD_ROW'
  | 'DELETE_ROW'
  | 'INSERT_ROW'
  | 'INSERT_COLUMN'
  | 'DELETE_COLUMN'
  | 'HIDE_ROW'
  | 'SHOW_ROW'
  | 'HIDE_COLUMN'
  | 'SHOW_COLUMN'
  | 'SET_ROW_HEIGHT'
  | 'SET_COLUMN_WIDTH'
  | 'FREEZE_PANES'
  | 'UNFREEZE_PANES'
  | 'MERGE_CELLS'
  | 'UNMERGE_CELLS'
  | 'CLEAR_CONTENT'
  | 'CLEAR_FORMAT'
  | 'CLEAR_ALL'
  | 'FORMAT_RANGE'
  | 'FILL_DOWN'
  | 'FILL_RIGHT'
  | 'CREATE_SHEET'
  | 'DELETE_SHEET'
  | 'RENAME_SHEET'
  | 'COPY_SHEET'
  | 'HIDE_SHEET'
  | 'SHOW_SHEET'
  | 'SET_SHEET_COLOR'
  | 'ADD_COMMENT'
  | 'DELETE_COMMENT'
  | 'WRITE_TABLE';

export interface SheetAction {
  type: SheetActionType;
  row?: number;
  col?: number;
  rowCount?: number;
  colCount?: number;
  value?: unknown;
  color?: string;
  formula?: string;
  data?: unknown[];
  count?: number;
  position?: 'above' | 'below' | 'left' | 'right' | 'before' | 'after';
  height?: number;
  width?: number;
  freezeRows?: number;
  freezeColumns?: number;
  mergeAcross?: boolean;
  format?: FormatSpec;
  endRow?: number;
  endCol?: number;
  sheetName?: string;
  newSheetName?: string;
  relativeTo?: string;
  comment?: string;
  headers?: string[];
  rows?: unknown[][];
}

export interface WorkbookContextPayload {
  activeSheet: string;
  sheets: string[];
}
