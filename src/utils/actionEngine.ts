import type { RowSnapshot } from '@/engine/rowSnapshot';
import { partitionActions } from '../engine/actionNormalizer';
import { richActionEngine, CreatedConditionalFormatId, CreatedChartId } from '../engine/actionEngine';
import type { FormatSnapshot } from '../engine/formatSnapshot';
import { SheetCreationOutcome } from '../engine/handlers/sheet.handler';
import { isOverwriteGuardError } from '../engine/overwriteGuard';
import {
  annotateDestOverwriteForCreatedSheets,
  pruneSpuriousAddSheets,
} from '../engine/overwriteGuard';
import { selectActionRanges } from '../engine/selectRanges';
import { SheetAction } from '../types/sheet-actions';
import { CellChange } from '../types/changeSet';
import { sanitizeActions, SanitizeResult } from './actionGuard';
import { probeSheetGuardStatesSafe } from '../engine/sheetGuardState';

/* global Excel */


/**
 * Guard the batch using live per-sheet facts rather than the batch's shape.
 *
 * This used to be `detectPopulateEmptySheet` — "does the batch contain both a
 * row-0 SET_CELL and an ADD_ROW?" — as a proxy for "is the target sheet
 * empty?". A multi-sheet build writes headers with SET_CELL and no ADD_ROW at
 * all, so the proxy said "populated", and the guard rewrote every header cell
 * in the batch into one row on one sheet. It also made the outcome depend on
 * whether an unrelated ADD_ROW happened to be present, so the same prompt
 * could succeed or fail run to run. See TASKS.md #137.
 *
 * Returns the full `SanitizeResult` — callers need `blocked`/`warnings` to
 * report a partial apply honestly instead of showing "Applied" over it.
 */
async function guardActions(actions: SheetAction[]): Promise<SanitizeResult> {
  const sheetStates = await probeSheetGuardStatesSafe(actions);
  return sanitizeActions(actions, undefined, { sheetStates });
}

export class ActionEngine {

  static async applyActions(actions: SheetAction[]): Promise<void> {
    const result = await this.applyActionsWithReport(actions);
    // Any error is a reportable failure, not just a total wipeout (see
    // applyActionsWithAudit in App.tsx for the same fix, and why).
    if (result.errors.length > 0) {
      throw new Error(result.errors.join('; '));
    }
  }

  static async applyActionsWithReport(actions: SheetAction[]): Promise<{
    applied: number;
    errors: string[];
    /** Actions the header guard dropped — never silently discard these. */
    guardBlocked?: SheetAction[];
    /** Human-readable notes on what the guard rewrote or dropped. */
    guardWarnings?: string[];
    createdConditionalFormatIds?: CreatedConditionalFormatId[];
    createdChartIds?: CreatedChartId[];
    sortedRangeChanges?: CellChange[];
    /** What each formatting action overwrote — Revert's data. TASKS.md #400. */
    formatSnapshots?: FormatSnapshot[];
    /** In-session undo data for row-level actions (P1.6). Memory only - never sent to the server. */
    rowSnapshots?: RowSnapshot[];
    /** ADD_SHEET/CREATE_SHEET outcomes where the live sheet name diverged from what was requested. */
    sheetNameMismatches?: SheetCreationOutcome[];
  }> {
    const errors: string[] = [];
    let applied = 0;
    let rowSnapshots: RowSnapshot[] | undefined;
    let createdConditionalFormatIds: CreatedConditionalFormatId[] | undefined;
    let createdChartIds: CreatedChartId[] | undefined;
    let sortedRangeChanges: CellChange[] | undefined;
    let formatSnapshots: FormatSnapshot[] | undefined;
    let sheetNameMismatches: SheetCreationOutcome[] | undefined;
    // Populated before the try so a throw still reports what the guard did.
    let guardBlocked: SheetAction[] = [];
    let guardWarnings: string[] = [];

    try {
      const guarded = await guardActions(actions);
      guardBlocked = guarded.blocked;
      guardWarnings = guarded.warnings;
      if (guardBlocked.length > 0) {
        // Spec 22's rule, applied to the guard: a batch that silently shrank must
        // not look like a clean apply. This used to be dropped on the floor —
        // 121 of 189 actions vanished and the UI still showed "Applied".
        console.warn(
          `[Cellix] Header guard dropped ${guardBlocked.length} of ${actions.length} action(s):`,
          guardWarnings,
        );
      }
      const safeInput = annotateDestOverwriteForCreatedSheets(
        pruneSpuriousAddSheets(guarded.actions),
      );
      const { rich, unsupported } = partitionActions(safeInput);
      if (unsupported.length > 0) {
        errors.push(
          `Unsupported action(s): ${unsupported.map((a) => a.type).join(', ')}. Unified engine only.`,
        );
      }
      if (!rich.length) {
        throw new Error(errors[0] ?? 'No actions to apply.');
      }

      const richResult = await richActionEngine.applyActions(rich);
      applied += richResult.applied;
      errors.push(...richResult.errors);
      createdConditionalFormatIds = richResult.createdConditionalFormatIds;
      createdChartIds = richResult.createdChartIds;
      sortedRangeChanges = richResult.sortedRangeChanges;
      formatSnapshots = richResult.formatSnapshots;
      rowSnapshots = richResult.rowSnapshots;
      sheetNameMismatches = richResult.sheetNameMismatches;
      if (richResult.errors.some((e) => e.includes('Write blocked:'))) {
        throw new Error(richResult.errors.find((e) => e.includes('Write blocked:')) ?? richResult.errors[0]);
      }
    } catch (error) {
      if (isOverwriteGuardError(error)) {
        throw error;
      }
      console.error('Failed to apply actions to spreadsheet:', error);
      throw new Error(
        `Spreadsheet update failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
    }

    return {
      applied,
      errors,
      ...(guardBlocked.length > 0 ? { guardBlocked } : {}),
      ...(guardWarnings.length > 0 ? { guardWarnings } : {}),
      ...(createdConditionalFormatIds ? { createdConditionalFormatIds } : {}),
      ...(createdChartIds ? { createdChartIds } : {}),
      ...(sortedRangeChanges ? { sortedRangeChanges } : {}),
      ...(formatSnapshots ? { formatSnapshots } : {}),
      ...(rowSnapshots ? { rowSnapshots } : {}),
      ...(sheetNameMismatches ? { sheetNameMismatches } : {}),
    };
  }

  /** Select (no fill) the ranges touched by these actions. */
  static async selectAppliedRanges(actions: SheetAction[]): Promise<void> {
    if (!actions.length) return;
    await Excel.run(async (ctx) => {
      await selectActionRanges(actions, ctx);
    });
  }

  static validateAction(action: SheetAction): boolean {
    switch (action.type) {
      case 'SET_CELL':
        return action.row !== undefined && action.col !== undefined && action.value !== undefined;
      case 'CLEAR_CELL':
      case 'HIGHLIGHT_CELL':
      case 'DELETE_COMMENT':
        return action.row !== undefined && action.col !== undefined;
      case 'SET_FORMULA':
        return action.row !== undefined && action.col !== undefined && action.formula !== undefined;
      case 'ADD_ROW':
        return Array.isArray(action.data);
      case 'DELETE_ROW':
      case 'INSERT_ROW':
      case 'HIDE_ROW':
      case 'UNHIDE_ROW':
      case 'SHOW_ROW':
      case 'SET_ROW_HEIGHT':
        return action.row !== undefined;
      case 'INSERT_COLUMN':
      case 'DELETE_COLUMN':
      case 'HIDE_COLUMN':
      case 'UNHIDE_COLUMN':
      case 'SHOW_COLUMN':
      case 'SET_COLUMN_WIDTH':
      case 'FILL_DOWN':
        return action.col !== undefined;
      case 'FILL_RIGHT':
        return action.row !== undefined && action.col !== undefined;
      case 'FORMAT_RANGE':
      case 'MERGE_CELLS':
      case 'UNMERGE_CELLS':
      case 'CLEAR_CONTENT':
      case 'CLEAR_FORMAT':
      case 'CLEAR_ALL':
      case 'ADD_COMMENT':
        return action.row !== undefined && action.col !== undefined;
      case 'FREEZE_PANES':
      case 'UNFREEZE_PANES':
      case 'PROTECT_SHEET':
      case 'UNPROTECT_SHEET':
      case 'CREATE_SHEET':
      case 'DELETE_SHEET':
      case 'RENAME_SHEET':
      case 'COPY_SHEET':
      case 'HIDE_SHEET':
      case 'SHOW_SHEET':
      case 'SET_SHEET_COLOR':
        return true;
      case 'WRITE_TABLE':
        return Array.isArray(action.headers) && action.headers.length > 0 && Array.isArray(action.rows);
      case 'BATCH_SET':
        return Boolean(action.sheetName) && Array.isArray(action.operations);
      case 'CREATE_TABLE':
        return Boolean(action.sheetName && action.range && action.tableName);
      case 'DEFINE_NAMED_RANGE':
        return Boolean(action.name && action.formula);
      case 'AUTOFIT_COLUMNS':
        return Boolean(action.sheetName);
      case 'CLARIFY':
        return Boolean(action.question);
      case 'CHECKPOINT':
        return Boolean(action.message);
      case 'ADD_SHEET':
        return Boolean(action.name ?? action.sheetName);
      case 'SET_ZOOM':
        return typeof action.zoomPercent === 'number';
      default:
        return false;
    }
  }
}

export type { SheetAction };
