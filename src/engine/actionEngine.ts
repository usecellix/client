import { RichAction } from '@/action.types';
import {
  handleAddRow,
  handleDeleteRow,
  handleInsertColumn,
  handleDeleteColumn,
} from './handlers/rowCol.handler';
import {
  handleSetCell,
  handleSetFormula,
  handleFillDown,
  handleBatchSet,
} from './handlers/cell.handler';
import {
  handleAddSheet,
  handleDeleteSheet,
  handleRenameSheet,
  handleCopySheet,
} from './handlers/sheet.handler';
import {
  handleCreateTable,
  handleDefineNamedRange,
  handleAutofitColumns,
} from './handlers/table.handler';
import { handleSortRange } from './handlers/sort.handler';
import { handleCopyFilteredRange, handleMoveRange, handleFormatMatchingRows, handleSetMatchingRows } from './handlers/range.handler';
import { handleAggregateTable } from './handlers/aggregate.handler';
import {
  handleAppendRow,
  handleAutoFill,
  handleClearRange,
  handleHighlightCell,
  handleInsertRow,
  handleMergeCells,
  handleWriteTable,
} from './handlers/misc.handler';
import { applyRichFormat } from './handlers/format.handler';
import { handleWorksheetAction } from './handlers/worksheet.handler';
import { handleCreateChart, handleUpdateChart } from './handlers/chart.handler';
import {
  annotateDestOverwriteForCreatedSheets,
  guardAgainstOverwrite,
  isOverwriteGuardError,
  OverwriteGuardError,
  pruneSpuriousAddSheets,
} from './overwriteGuard';
import { selectActionRanges } from './selectRanges';
import { resolveWorksheet } from './sheetResolve';

/* global Excel */

export class RichActionEngine {
  async applyActions(actions: RichAction[]): Promise<{ applied: number; errors: string[] }> {
    const errors: string[] = [];
    let applied = 0;
    const prepared = annotateDestOverwriteForCreatedSheets(
      pruneSpuriousAddSheets(actions),
    );

    if (prepared.length === 0) {
      return { applied, errors };
    }

    try {
      await Excel.run(async (ctx) => {
        for (const action of prepared) {
          try {
            await this.dispatch(action, ctx);
            applied += 1;
          } catch (err: unknown) {
            // Overwrite guard is a hard stop — do not continue writing after a block.
            if (isOverwriteGuardError(err)) {
              throw err;
            }
            const message = err instanceof Error ? err.message : String(err);
            errors.push(`${action.type}: ${message}`);
            console.error(`RichActionEngine error on ${action.type}:`, err);
          }
        }

        // Mouse-select edited area (no fill colors) after successful writes.
        if (applied > 0) {
          try {
            await selectActionRanges(prepared, ctx);
          } catch (selectErr) {
            console.warn('[Cellix] Failed to select applied ranges:', selectErr);
          }
        }

        await ctx.sync();
      });
    } catch (err: unknown) {
      if (isOverwriteGuardError(err)) {
        const message = err instanceof Error ? err.message : String(err);
        errors.push(message);
        throw err instanceof OverwriteGuardError
          ? err
          : new OverwriteGuardError({
              message,
              targetRange: (err as OverwriteGuardError).targetRange ?? 'unknown',
              sampleExistingValues:
                (err as OverwriteGuardError).sampleExistingValues ?? [],
            });
      }
      const message = err instanceof Error ? err.message : String(err);
      errors.push(message);
    }

    return { applied, errors };
  }

  private async dispatch(action: RichAction, ctx: Excel.RequestContext): Promise<void> {
    // Last line of defense: never silently overwrite occupied cells.
    await guardAgainstOverwrite(action, ctx);

    if (await handleWorksheetAction(action, ctx)) {
      return;
    }

    switch (action.type) {
      case 'ADD_ROW':
        return handleAddRow(action, ctx);
      case 'APPEND_ROW':
        return handleAppendRow(action, ctx);
      case 'INSERT_ROW':
        return handleInsertRow(action, ctx);
      case 'SET_CELL':
        return handleSetCell(action, ctx);
      case 'FORMAT_RANGE': {
        const sheet = resolveWorksheet(ctx, action.sheetName);
        const range = sheet.getRange(action.range);
        applyRichFormat(range, action.format);
        return;
      }
      case 'SET_FORMULA':
        return handleSetFormula(action, ctx);
      case 'FILL_DOWN':
        return handleFillDown(action, ctx);
      case 'AUTO_FILL':
        return handleAutoFill(action, ctx);
      case 'BATCH_SET':
        return handleBatchSet(action, ctx);
      case 'DELETE_ROW':
        return handleDeleteRow(action, ctx);
      case 'INSERT_COLUMN':
        return handleInsertColumn(action, ctx);
      case 'DELETE_COLUMN':
        return handleDeleteColumn(action, ctx);
      case 'ADD_SHEET':
        return handleAddSheet(action, ctx);
      case 'DELETE_SHEET':
        return handleDeleteSheet(action, ctx);
      case 'RENAME_SHEET':
        return handleRenameSheet(action, ctx);
      case 'COPY_SHEET':
        return handleCopySheet(action, ctx);
      case 'CREATE_TABLE':
        return handleCreateTable(action, ctx);
      case 'CREATE_CHART': {
        const result = await handleCreateChart(action, ctx);
        if (result.chartId && !(action as { chartId?: string }).chartId) {
          (action as { chartId?: string }).chartId = result.chartId;
        }
        return;
      }
      case 'UPDATE_CHART':
        return handleUpdateChart(action, ctx);
      case 'AGGREGATE_TABLE':
        await handleAggregateTable(action, ctx);
        return;
      case 'DEFINE_NAMED_RANGE':
        return handleDefineNamedRange(action, ctx);
      case 'AUTOFIT_COLUMNS':
        return handleAutofitColumns(action, ctx);
      case 'WRITE_TABLE':
        return handleWriteTable(action, ctx);
      case 'HIGHLIGHT_CELL':
        return handleHighlightCell(action, ctx);
      case 'MERGE_CELLS':
        return handleMergeCells(action, ctx);
      case 'CLEAR_RANGE':
        return handleClearRange(action, ctx);
      case 'SORT_RANGE':
        return handleSortRange(action, ctx);
      case 'COPY_FILTERED_RANGE':
        await handleCopyFilteredRange(action, ctx);
        return;
      case 'FORMAT_MATCHING_ROWS':
        await handleFormatMatchingRows(action, ctx);
        return;
      case 'SET_MATCHING_ROWS':
        await handleSetMatchingRows(action, ctx);
        return;
      case 'MOVE_RANGE':
        await handleMoveRange(action, ctx);
        return;
      case 'CLARIFY':
      case 'CHECKPOINT':
        return;
      default:
        console.warn('Unknown rich action type:', (action as RichAction).type);
    }
  }
}

export const richActionEngine = new RichActionEngine();
