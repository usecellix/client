import { SheetActionType } from './sheet-actions';

/**
 * Every SheetActionType must have an entry here — adding a new wire-level
 * action type without one fails the build. This is the frontend's mirror of
 * cellix_backend/src/excel-ai/types/action-catalog.ts's ACTION_CATALOG, one
 * layer up: that file guarantees the backend's own union stays internally
 * consistent; this one lets a cross-repo test confirm the frontend at least
 * recognizes every wire type the backend's catalog can advertise or emit,
 * before actionNormalizer.ts / legacyConverter.ts narrow it down to a
 * dispatchable RichAction. See TASKS.md #5.
 */
const SHEET_ACTION_CATALOG: Record<SheetActionType, true> = {
  SET_CELL: true,
  CLEAR_CELL: true,
  HIGHLIGHT_CELL: true,
  SET_FORMULA: true,
  ADD_ROW: true,
  DELETE_ROW: true,
  INSERT_ROW: true,
  INSERT_COLUMN: true,
  DELETE_COLUMN: true,
  HIDE_ROW: true,
  UNHIDE_ROW: true,
  SHOW_ROW: true,
  HIDE_COLUMN: true,
  UNHIDE_COLUMN: true,
  SHOW_COLUMN: true,
  SET_ROW_HEIGHT: true,
  SET_COLUMN_WIDTH: true,
  FREEZE_PANES: true,
  UNFREEZE_PANES: true,
  AUTO_FILTER: true,
  CLEAR_RANGE: true,
  APPEND_ROW: true,
  SET_ZOOM: true,
  PROTECT_SHEET: true,
  UNPROTECT_SHEET: true,
  MERGE_CELLS: true,
  UNMERGE_CELLS: true,
  CLEAR_CONTENT: true,
  CLEAR_FORMAT: true,
  CLEAR_ALL: true,
  FORMAT_RANGE: true,
  FILL_DOWN: true,
  FILL_RIGHT: true,
  CREATE_SHEET: true,
  DELETE_SHEET: true,
  RENAME_SHEET: true,
  COPY_SHEET: true,
  HIDE_SHEET: true,
  SHOW_SHEET: true,
  SET_SHEET_COLOR: true,
  ADD_COMMENT: true,
  DELETE_COMMENT: true,
  WRITE_TABLE: true,
  BATCH_SET: true,
  CREATE_TABLE: true,
  DELETE_TABLE: true,
  CREATE_CHART: true,
  DEFINE_NAMED_RANGE: true,
  AUTOFIT_COLUMNS: true,
  CLARIFY: true,
  CHECKPOINT: true,
  ADD_SHEET: true,
  SORT_RANGE: true,
  COPY_FILTERED_RANGE: true,
  FORMAT_MATCHING_ROWS: true,
  SET_MATCHING_ROWS: true,
  MOVE_RANGE: true,
  AGGREGATE_TABLE: true,
  UPDATE_CHART: true,
  DELETE_CHART: true,
  CONDITIONAL_FORMAT: true,
  DELETE_CONDITIONAL_FORMAT: true,
};

/** Every wire-level action type the frontend recognizes, derived from the catalog above. */
export const ALL_FRONTEND_SHEET_ACTION_TYPES = Object.keys(
  SHEET_ACTION_CATALOG,
) as SheetActionType[];
