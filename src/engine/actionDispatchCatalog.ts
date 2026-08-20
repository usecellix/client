import { RichAction } from '@/action.types';

/**
 * Every RichAction type must have an entry here — adding a new type to
 * shared/action.types.ts without one fails the build. Mirrors
 * cellix_backend/src/excel-ai/types/action-catalog.ts's pattern, applied to
 * the frontend's own dispatch completeness: this guards "actionEngine.ts's
 * switch (or handleWorksheetAction) actually handles every canonical action,"
 * independent of the wire-level check in types/sheetActionCatalog.ts. See
 * TASKS.md #5.
 */
const DISPATCH_CATALOG: Record<RichAction['type'], true> = {
  SET_CELL: true,
  SET_FORMULA: true,
  FORMAT_RANGE: true,
  FILL_DOWN: true,
  BATCH_SET: true,
  ADD_ROW: true,
  APPEND_ROW: true,
  INSERT_ROW: true,
  DELETE_ROW: true,
  INSERT_COLUMN: true,
  DELETE_COLUMN: true,
  AUTO_FILL: true,
  WRITE_TABLE: true,
  HIGHLIGHT_CELL: true,
  MERGE_CELLS: true,
  CLEAR_RANGE: true,
  ADD_SHEET: true,
  DELETE_SHEET: true,
  RENAME_SHEET: true,
  COPY_SHEET: true,
  CREATE_TABLE: true,
  DELETE_TABLE: true,
  CREATE_CHART: true,
  UPDATE_CHART: true,
  DELETE_CHART: true,
  AGGREGATE_TABLE: true,
  DEFINE_NAMED_RANGE: true,
  AUTOFIT_COLUMNS: true,
  SORT_RANGE: true,
  CLARIFY: true,
  CHECKPOINT: true,
  COPY_FILTERED_RANGE: true,
  FORMAT_MATCHING_ROWS: true,
  SET_MATCHING_ROWS: true,
  MOVE_RANGE: true,
  HIDE_ROW: true,
  UNHIDE_ROW: true,
  HIDE_COLUMN: true,
  UNHIDE_COLUMN: true,
  SET_ROW_HEIGHT: true,
  SET_COLUMN_WIDTH: true,
  FREEZE_PANES: true,
  UNFREEZE_PANES: true,
  AUTO_FILTER: true,
  SET_ZOOM: true,
  PROTECT_SHEET: true,
  UNPROTECT_SHEET: true,
  UNMERGE_CELLS: true,
  HIDE_SHEET: true,
  SHOW_SHEET: true,
  SET_SHEET_COLOR: true,
  ADD_COMMENT: true,
  DELETE_COMMENT: true,
  CONDITIONAL_FORMAT: true,
  DELETE_CONDITIONAL_FORMAT: true,
};

/** Every RichAction type the frontend can dispatch, derived from the catalog above. */
export const ALL_RICH_ACTION_TYPES = Object.keys(DISPATCH_CATALOG) as RichAction['type'][];
