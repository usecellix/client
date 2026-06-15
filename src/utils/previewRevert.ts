import { SheetAction } from '@/types/sheet-actions';
import { CellChange } from '@/types/changeSet';
import { parseCellAddress } from '@/engine/addressUtils';

/** Applied on preview start (visible immediately). DELETE_SHEET is excluded — apply only on Accept. */
export const STRUCTURAL_PREVIEW_ACTION_TYPES = new Set<SheetAction['type']>([
  'ADD_SHEET',
  'CREATE_SHEET',
  'COPY_SHEET',
  'RENAME_SHEET',
  'ADD_ROW',
  'DELETE_ROW',
  'INSERT_ROW',
  'INSERT_COLUMN',
  'DELETE_COLUMN',
  'WRITE_TABLE',
  'SORT_RANGE',
  'BATCH_SET',
  'CREATE_TABLE',
  'DEFINE_NAMED_RANGE',
]);

/** Never applied during preview — only on Accept (Reject leaves workbook unchanged). */
export const DEFERRED_PREVIEW_ACTION_TYPES = new Set<SheetAction['type']>(['DELETE_SHEET']);

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
      });
    } else {
      actions.push({
        type: 'SET_CELL',
        sheetName: change.sheet,
        row: parsed.row,
        col: parsed.col,
        value: change.before as string | number | boolean | null,
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

  if (options.deferredApplied) {
    revert.push(...buildCellRevertActions(changes));
    revert.push(...buildStructuralRevertActions(deferred));
  }

  if (options.structuralApplied) {
    revert.push(...buildStructuralRevertActions(structural));
  }

  return revert;
}
