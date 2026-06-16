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

/* global Excel */

export class RichActionEngine {
  async applyActions(actions: RichAction[]): Promise<{ applied: number; errors: string[] }> {
    const errors: string[] = [];
    let applied = 0;

    if (actions.length === 0) {
      return { applied, errors };
    }

    try {
      await Excel.run(async (ctx) => {
        for (const action of actions) {
          try {
            await this.dispatch(action, ctx);
            applied += 1;
          } catch (err: unknown) {
            const message = err instanceof Error ? err.message : String(err);
            errors.push(`${action.type}: ${message}`);
            console.error(`RichActionEngine error on ${action.type}:`, err);
          }
        }
        await ctx.sync();
      });
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      errors.push(message);
    }

    return { applied, errors };
  }

  private async dispatch(action: RichAction, ctx: Excel.RequestContext): Promise<void> {
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
        const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
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
      case 'CLARIFY':
      case 'CHECKPOINT':
        return;
      default:
        console.warn('Unknown rich action type:', (action as RichAction).type);
    }
  }
}

export const richActionEngine = new RichActionEngine();
