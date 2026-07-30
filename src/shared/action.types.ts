export type CellValue = string | number | boolean | null;

export type ChartColorScheme =
  | 'default'
  | 'blue'
  | 'grey'
  | 'blueGrey'
  | 'green'
  | 'red'
  | 'orange'
  | 'purple'
  | 'yellow';

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

/** Rich FormatSpec — format.handler reads borders.edges */
export interface FormatSpec {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  clearFill?: boolean;
  numberFormat?: string;
  horizontalAlignment?: 'left' | 'center' | 'right';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  borders?: {
    style: 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'none';
    color?: string;
    edges: ('top' | 'bottom' | 'left' | 'right' | 'all' | 'outer' | 'inner')[];
  };
}

export interface BatchSetOperation {
  address: string;
  value?: CellValue;
  formula?: string;
  format?: FormatSpec;
}

export type InsertColumnPosition = 'afterLastColumn' | { afterColumn: string };

export type AggregateFn = 'sum' | 'count' | 'average' | 'max' | 'min';

export type GroupByTransform =
  | 'none'
  | 'month'
  | 'year'
  | 'monthYear'
  | 'weekday'
  | 'quarter';

export interface AggregateSpec {
  column: string;
  fn: AggregateFn;
  outputLabel: string;
}

export interface SetCellAction {
  type: 'SET_CELL';
  sheetName: string;
  address: string;
  value: CellValue;
  format?: FormatSpec;
  explicitOverwriteConfirmed?: boolean;
}

export interface SetFormulaAction {
  type: 'SET_FORMULA';
  sheetName: string;
  address: string;
  formula: string;
  format?: FormatSpec;
  explicitOverwriteConfirmed?: boolean;
}

export interface FormatRangeAction {
  type: 'FORMAT_RANGE';
  sheetName: string;
  range: string;
  format: FormatSpec;
}

export interface FillDownAction {
  type: 'FILL_DOWN';
  sheetName: string;
  sourceRange: string;
  targetRange: string;
  explicitOverwriteConfirmed?: boolean;
}

export interface BatchSetAction {
  type: 'BATCH_SET';
  sheetName: string;
  operations: BatchSetOperation[];
  explicitOverwriteConfirmed?: boolean;
}

export interface AddRowAction {
  type: 'ADD_ROW';
  sheetName: string;
  afterRow: number;
  values: CellValue[];
  copyFormatFromRow?: number;
}

export interface AppendRowAction {
  type: 'APPEND_ROW';
  sheetName: string;
  values: CellValue[];
  explicitOverwriteConfirmed?: boolean;
}

export interface InsertRowAction {
  type: 'INSERT_ROW';
  sheetName: string;
  row: number;
  count?: number;
  position?: 'above' | 'below';
}

export interface DeleteRowAction {
  type: 'DELETE_ROW';
  sheetName: string;
  rows: number[];
}

export interface InsertColumnAction {
  type: 'INSERT_COLUMN';
  sheetName: string;
  beforeColumn?: string;
  count?: number;
  copyFormatFromColumn?: string;
  columnName?: string;
  position?: InsertColumnPosition;
  afterColumn?: string;
  formula?: string;
  explicitOverwriteConfirmed?: boolean;
}

export interface DeleteColumnAction {
  type: 'DELETE_COLUMN';
  sheetName: string;
  columns: string[];
}

export interface AddSheetAction {
  type: 'ADD_SHEET';
  name: string;
  position?: number;
  copyFrom?: string;
}

export interface DeleteSheetAction {
  type: 'DELETE_SHEET';
  sheetName: string;
}

export interface RenameSheetAction {
  type: 'RENAME_SHEET';
  oldName: string;
  newName: string;
}

export interface CopySheetAction {
  type: 'COPY_SHEET';
  sourceName: string;
  newName: string;
  position?: number;
}

export interface CreateTableAction {
  type: 'CREATE_TABLE';
  sheetName: string;
  range: string;
  tableName: string;
  hasHeaders: boolean;
  style?: string;
}

export interface DefineNamedRangeAction {
  type: 'DEFINE_NAMED_RANGE';
  name: string;
  formula: string;
  comment?: string;
}

export interface AutoFitColumnsAction {
  type: 'AUTOFIT_COLUMNS';
  sheetName: string;
  columns?: string[];
}

export interface WriteTableAction {
  type: 'WRITE_TABLE';
  sheetName: string;
  headers: CellValue[];
  rows: CellValue[][];
  explicitOverwriteConfirmed?: boolean;
}

export interface HighlightCellAction {
  type: 'HIGHLIGHT_CELL';
  sheetName: string;
  address: string;
  color?: string;
}

export interface MergeCellsAction {
  type: 'MERGE_CELLS';
  sheetName: string;
  range: string;
}

export interface ClearRangeAction {
  type: 'CLEAR_RANGE';
  sheetName: string;
  range: string;
  mode: 'contents' | 'formats' | 'all';
}

export interface AutoFillAction {
  type: 'AUTO_FILL';
  sheetName: string;
  startAddress: string;
  direction: 'down' | 'right';
  endRow?: number;
  endCol?: number;
  explicitOverwriteConfirmed?: boolean;
}

export interface SortRangeAction {
  type: 'SORT_RANGE';
  sheetName: string;
  range: string;
  key: number;
  ascending: boolean;
  hasHeaders: boolean;
  columnName?: string;
}

export interface CreateChartAction {
  type: 'CREATE_CHART';
  sheetName: string;
  sourceSheetName: string;
  sourceRange: string;
  chartType: string;
  title?: string;
  startCell?: string;
  endCell?: string;
  destCell?: string;
  colorScheme?: ChartColorScheme;
  chartId?: string;
}

export interface UpdateChartAction {
  type: 'UPDATE_CHART';
  sheetName: string;
  chartId: string;
  chartType?: string;
  colorScheme?: ChartColorScheme;
}

export interface CopyFilteredRangeAction {
  type: 'COPY_FILTERED_RANGE';
  sourceSheet: string;
  sourceRange: string;
  hasHeaders: boolean;
  destSheet: string;
  destStartCell: string;
  filter?: RangeFilterSpec;
  mode: 'copy' | 'move';
  explicitOverwriteConfirmed?: boolean;
}

export interface MoveRangeAction {
  type: 'MOVE_RANGE';
  sourceSheet: string;
  sourceRange: string;
  destSheet: string;
  destStartCell: string;
  explicitOverwriteConfirmed?: boolean;
}

export interface FormatMatchingRowsAction {
  type: 'FORMAT_MATCHING_ROWS';
  sheetName: string;
  range: string;
  hasHeaders: boolean;
  filter: RangeFilterSpec;
  format: FormatSpec;
}

export interface AggregateTableAction {
  type: 'AGGREGATE_TABLE';
  sourceSheet: string;
  sourceRange: string;
  groupByColumn: string;
  groupByTransform?: GroupByTransform;
  aggregations: AggregateSpec[];
  sortBy?: { column: string; direction: 'asc' | 'desc' };
  topN?: number;
  destSheet: string;
  destStartCell: string;
  hasHeaders: boolean;
  explicitOverwriteConfirmed?: boolean;
}

export interface ClarifyAction {
  type: 'CLARIFY';
  question: string;
  options?: string[];
}

export interface CheckpointAction {
  type: 'CHECKPOINT';
  message: string;
}

export interface HideRowAction {
  type: 'HIDE_ROW' | 'UNHIDE_ROW' | 'SHOW_ROW';
  sheetName: string;
  row: number;
  rowCount?: number;
}

export interface HideColumnAction {
  type: 'HIDE_COLUMN' | 'UNHIDE_COLUMN' | 'SHOW_COLUMN';
  sheetName: string;
  col: number;
  colCount?: number;
}

export interface SetRowHeightAction {
  type: 'SET_ROW_HEIGHT';
  sheetName: string;
  row: number;
  height: number;
}

export interface SetColumnWidthAction {
  type: 'SET_COLUMN_WIDTH';
  sheetName: string;
  col: number;
  width: number;
}

export interface FreezePanesAction {
  type: 'FREEZE_PANES';
  sheetName: string;
  freezeRows?: number;
  freezeColumns?: number;
}

export interface UnfreezePanesAction {
  type: 'UNFREEZE_PANES';
  sheetName: string;
}

export interface SetZoomAction {
  type: 'SET_ZOOM';
  sheetName: string;
  zoomPercent: number;
}

export interface ProtectSheetAction {
  type: 'PROTECT_SHEET' | 'UNPROTECT_SHEET';
  sheetName: string;
}

export interface UnmergeCellsAction {
  type: 'UNMERGE_CELLS';
  sheetName: string;
  range: string;
}

export interface HideSheetAction {
  type: 'HIDE_SHEET' | 'SHOW_SHEET';
  sheetName: string;
}

export interface SetSheetColorAction {
  type: 'SET_SHEET_COLOR';
  sheetName: string;
  color: string;
}

export interface AddCommentAction {
  type: 'ADD_COMMENT';
  sheetName: string;
  address: string;
  comment: string;
}

export interface DeleteCommentAction {
  type: 'DELETE_COMMENT';
  sheetName: string;
  address: string;
}

export type RichAction =
  | SetCellAction
  | SetFormulaAction
  | FormatRangeAction
  | FillDownAction
  | BatchSetAction
  | AddRowAction
  | AppendRowAction
  | InsertRowAction
  | DeleteRowAction
  | InsertColumnAction
  | DeleteColumnAction
  | AddSheetAction
  | DeleteSheetAction
  | RenameSheetAction
  | CopySheetAction
  | CreateTableAction
  | DefineNamedRangeAction
  | AutoFitColumnsAction
  | WriteTableAction
  | HighlightCellAction
  | MergeCellsAction
  | ClearRangeAction
  | AutoFillAction
  | SortRangeAction
  | CreateChartAction
  | UpdateChartAction
  | CopyFilteredRangeAction
  | MoveRangeAction
  | FormatMatchingRowsAction
  | AggregateTableAction
  | ClarifyAction
  | CheckpointAction
  | HideRowAction
  | HideColumnAction
  | SetRowHeightAction
  | SetColumnWidthAction
  | FreezePanesAction
  | UnfreezePanesAction
  | SetZoomAction
  | ProtectSheetAction
  | UnmergeCellsAction
  | HideSheetAction
  | SetSheetColorAction
  | AddCommentAction
  | DeleteCommentAction;
