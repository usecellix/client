import { SheetAction } from '@/types/sheet-actions';
import { CellChange } from '@/types/changeSet';
import { parseCellAddress } from '@/engine/addressUtils';

/**
 * Applied on preview start (visible immediately).
 * Sheet creates/renames only — data mutations wait for Accept so Reject is clean
 * and Accept cannot double-apply INSERT_COLUMN / ADD_ROW.
 */
export const STRUCTURAL_PREVIEW_ACTION_TYPES = new Set<SheetAction['type']>([
  'ADD_SHEET',
  'CREATE_SHEET',
  'COPY_SHEET',
  'RENAME_SHEET',
  'CREATE_TABLE',
  'DEFINE_NAMED_RANGE',
]);

/**
 * Never applied during preview — only on Accept (Reject leaves workbook unchanged).
 * Includes row/column inserts so new columns/rows appear only after Accept.
 */
export const DEFERRED_PREVIEW_ACTION_TYPES = new Set<SheetAction['type']>([
  'DELETE_SHEET',
  'SORT_RANGE',
  'COPY_FILTERED_RANGE',
  'FORMAT_MATCHING_ROWS',
  'MOVE_RANGE',
  'CREATE_CHART',
  'UPDATE_CHART',
  'AGGREGATE_TABLE',
  'INSERT_COLUMN',
  'DELETE_COLUMN',
  'ADD_ROW',
  'INSERT_ROW',
  'DELETE_ROW',
  'WRITE_TABLE',
  'BATCH_SET',
  // Cell writes also wait for Accept — preview shows the diff list only.
  'SET_CELL',
  'SET_FORMULA',
  'CLEAR_CELL',
  'CLEAR_CONTENT',
  'CLEAR_FORMAT',
  'CLEAR_ALL',
  'FILL_DOWN',
  'FILL_RIGHT',
  'FORMAT_RANGE',
  'HIGHLIGHT_CELL',
  'MERGE_CELLS',
]);

export function partitionPreviewActions(actions: SheetAction[]): {
  structural: SheetAction[];
  deferred: SheetAction[];
} {
  const structural: SheetAction[] = [];
  const deferred: SheetAction[] = [];

  for (const action of actions) {
    if (DEFERRED_PREVIEW_ACTION_TYPES.has(action.type)) {
      deferred.push(action);
    } else if (STRUCTURAL_PREVIEW_ACTION_TYPES.has(action.type)) {
      structural.push(action);
    } else {
      deferred.push(action);
    }
  }

  return { structural, deferred };
}

function sheetNameFromAdd(action: SheetAction): string {
  return String(action.name ?? action.sheetName ?? '');
}

/** Inverse structural actions in reverse order (for reject after preview apply). */
export function buildStructuralRevertActions(actions: SheetAction[]): SheetAction[] {
  const inverse: SheetAction[] = [];

  for (let i = actions.length - 1; i >= 0; i -= 1) {
    const action = actions[i];
    switch (action.type) {
      case 'ADD_SHEET':
      case 'CREATE_SHEET': {
        const name = sheetNameFromAdd(action);
        if (name) inverse.push({ type: 'DELETE_SHEET', sheetName: name });
        break;
      }
      case 'COPY_SHEET': {
        const name = String(action.newName ?? action.newSheetName ?? '');
        if (name) inverse.push({ type: 'DELETE_SHEET', sheetName: name });
        break;
      }
      case 'RENAME_SHEET':
        inverse.push({
          type: 'RENAME_SHEET',
          sheetName: action.newName ?? action.newSheetName ?? action.sheetName,
          newSheetName: action.oldName ?? action.sheetName ?? '',
        });
        break;
      default:
        break;
    }
  }

  return inverse;
}

/** Restore cell values from audited change list (reject before deferred cell apply). */
export function buildCellRevertActions(changes: CellChange[]): SheetAction[] {
  const actions: SheetAction[] = [];

  for (const change of changes) {
    const parsed = parseCellAddress(change.cell);
    if (!parsed) continue;

    if (change.formula && String(change.before ?? '').startsWith('=')) {
      actions.push({
        type: 'SET_FORMULA',
        sheetName: change.sheet,
        row: parsed.row,
        col: parsed.col,
        formula: String(change.before),
        explicitOverwriteConfirmed: true,
      });
    } else {
      actions.push({
        type: 'SET_CELL',
        sheetName: change.sheet,
        row: parsed.row,
        col: parsed.col,
        value: change.before as string | number | boolean | null,
        explicitOverwriteConfirmed: true,
      });
    }
  }

  return actions;
}

export function buildPreviewRejectActions(
  actions: SheetAction[],
  changes: CellChange[] = [],
  options: { structuralApplied?: boolean; deferredApplied?: boolean } = {},
): SheetAction[] {
  const { structural, deferred } = partitionPreviewActions(actions);
  const revert: SheetAction[] = [];

  // Cell writes may be applied during preview (early deferred) or on Accept.
  // SORT_RANGE is hard-deferred (never applied in preview), so no sort undo is needed.
  if (options.deferredApplied || options.structuralApplied) {
    revert.push(...buildCellRevertActions(changes));
  }

  if (options.deferredApplied) {
    revert.push(...buildStructuralRevertActions(deferred));
  }

  if (options.structuralApplied) {
    revert.push(...buildStructuralRevertActions(structural));
  }

  return revert;
}
