export interface FormatSpec {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  clearFill?: boolean;
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
  | 'UNHIDE_ROW'
  | 'SHOW_ROW'
  | 'HIDE_COLUMN'
  | 'UNHIDE_COLUMN'
  | 'SHOW_COLUMN'
  | 'SET_ROW_HEIGHT'
  | 'SET_COLUMN_WIDTH'
  | 'FREEZE_PANES'
  | 'UNFREEZE_PANES'
  | 'AUTO_FILTER'
  // Rich-only forms produced by legacyConverter and the local shortcuts.
  | 'CLEAR_RANGE'
  | 'APPEND_ROW'
  | 'SET_ZOOM'
  | 'PROTECT_SHEET'
  | 'UNPROTECT_SHEET'
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
  | 'WRITE_TABLE'
  | 'BATCH_SET'
  | 'CREATE_TABLE'
  | 'CREATE_CHART'
  | 'DEFINE_NAMED_RANGE'
  | 'AUTOFIT_COLUMNS'
  | 'CLARIFY'
  | 'CHECKPOINT'
  | 'ADD_SHEET'
  | 'SORT_RANGE'
  | 'COPY_FILTERED_RANGE'
  | 'FORMAT_MATCHING_ROWS'
  | 'SET_MATCHING_ROWS'
  | 'MOVE_RANGE'
  | 'AGGREGATE_TABLE'
  | 'UPDATE_CHART';

export type RangeFilterOperator =
  | 'equals'
  | 'contains'
  | 'greaterThan'
  | 'lessThan'
  | 'notEquals'
  | 'lengthEquals'
  | 'lengthNotEquals'
  | 'matchesRegex'
  | 'notMatchesRegex';

export interface RangeFilterSpec {
  column: string;
  operator: RangeFilterOperator;
  value: string | number;
}

export interface BatchSetOperation {
  address: string;
  value?: unknown;
  formula?: string;
  format?: FormatSpec;
}

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
  position?: 'above' | 'below' | 'left' | 'right' | 'before' | 'after' | 'afterLastColumn';
  height?: number;
  width?: number;
  freezeRows?: number;
  freezeColumns?: number;
  zoomPercent?: number;
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
  /** Rich action: Excel 1-based row after which to insert */
  afterRow?: number;
  values?: unknown[];
  copyFormatFromRow?: number;
  address?: string;
  range?: string;
  sourceRange?: string;
  targetRange?: string;
  sourceSheetName?: string;
  chartType?: string;
  title?: string;
  startCell?: string;
  endCell?: string;
  operations?: BatchSetOperation[];
  /** Rich DELETE_ROW: 1-based row numbers (when type is DELETE_ROW) */
  rowNumbers?: number[];
  beforeColumn?: string;
  /** INSERT_COLUMN semantic: insert after this header name */
  afterColumn?: string;
  columns?: string[];
  oldName?: string;
  /** RENAME_SHEET / COPY_SHEET target name — the backend payload sends this. */
  newName?: string;
  sourceName?: string;
  name?: string;
  tableName?: string;
  hasHeaders?: boolean;
  style?: string;
  question?: string;
  options?: string[];
  message?: string;
  copyFrom?: string;
  key?: number;
  ascending?: boolean;
  columnName?: string;
  /** SET_MATCHING_ROWS — column to write into */
  targetColumn?: string;
  /** COPY_FILTERED_RANGE / MOVE_RANGE / FORMAT_MATCHING_ROWS / SET_MATCHING_ROWS */
  sourceSheet?: string;
  destSheet?: string;
  destStartCell?: string;
  filter?: RangeFilterSpec;
  /** MOVE_RANGE / COPY_FILTERED_RANGE use copy|move; CLEAR_RANGE uses the clear modes. */
  mode?: 'copy' | 'move' | 'contents' | 'formats' | 'all';
  groupByColumn?: string;
  groupByTransform?: 'none' | 'month' | 'year' | 'monthYear' | 'weekday' | 'quarter';
  aggregations?: Array<{
    column: string;
    /** 'first' passes through a label column 1:1 with the group key (e.g. Supplier Name alongside a GSTIN group-by) — not a numeric reduction. */
    fn: 'sum' | 'count' | 'average' | 'max' | 'min' | 'first';
    outputLabel: string;
  }>;
  sortBy?: { column: string; direction: 'asc' | 'desc' };
  topN?: number;
  destCell?: string;
  colorScheme?: 'default' | 'blue' | 'grey' | 'blueGrey' | 'green' | 'red' | 'orange' | 'purple' | 'yellow';
  chartId?: string;
  /**
   * When true, allow writing over non-empty cells.
   * Only set for unambiguously replace/clear intents — never inferred by the Executor.
   */
  explicitOverwriteConfirmed?: boolean;
}

export interface WorkbookContextPayload {
  activeSheet: string;
  sheets: string[];
}
