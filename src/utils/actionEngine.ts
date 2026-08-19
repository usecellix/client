import { partitionActions } from '../engine/actionNormalizer';
import { richActionEngine } from '../engine/actionEngine';
import { isOverwriteGuardError } from '../engine/overwriteGuard';
import {
  annotateDestOverwriteForCreatedSheets,
  pruneSpuriousAddSheets,
} from '../engine/overwriteGuard';
import { selectActionRanges } from '../engine/selectRanges';
import { SheetAction, SheetActionType } from '../types/sheet-actions';
import { sanitizeActions } from './actionGuard';

/* global Excel */

type PreviewRange = {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
  values: unknown[][];
  formulas: unknown[][];
  fillColor: string | null;
  // Office.js widens this to the enum plus its string-literal aliases; mirror that
  // so a captured pattern can be written straight back on revert.
  fillPattern: Excel.RangeFill['pattern'];
};


function detectPopulateEmptySheet(actions: SheetAction[]): boolean {
  const hasHeaderCells = actions.some((a) => a.type === 'SET_CELL' && a.row === 0);
  const hasAddRows = actions.some((a) => a.type === 'ADD_ROW');
  return hasHeaderCells && hasAddRows;
}

function guardActions(actions: SheetAction[], sheetIsEmpty = false): SheetAction[] {
  const isEmpty = sheetIsEmpty || detectPopulateEmptySheet(actions);
  const layout = isEmpty
    ? { headerRow: 0, nextDataRow: 0, dataRowCount: 0, columnCount: 1, headers: [], isEmpty: true }
    : undefined;
  const { actions: safe } = sanitizeActions(actions, layout);
  return safe;
}

export class ActionEngine {
  private static previewRanges: PreviewRange[] = [];

  static async applyActions(actions: SheetAction[]): Promise<void> {
    const result = await this.applyActionsWithReport(actions);
    if (result.errors.length > 0 && result.applied === 0) {
      throw new Error(result.errors.join('; '));
    }
  }

  static async applyActionsWithReport(actions: SheetAction[]): Promise<{
    applied: number;
    errors: string[];
  }> {
    const errors: string[] = [];
    let applied = 0;

    try {
      const safeInput = annotateDestOverwriteForCreatedSheets(
        pruneSpuriousAddSheets(guardActions(actions)),
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

      if (this.previewRanges.length) {
        await this.commitPreview();
      }

      const richResult = await richActionEngine.applyActions(rich);
      applied += richResult.applied;
      errors.push(...richResult.errors);
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

    return { applied, errors };
  }

  /** Select (no fill) the ranges touched by these actions. */
  static async selectAppliedRanges(actions: SheetAction[]): Promise<void> {
    if (!actions.length) return;
    await Excel.run(async (ctx) => {
      await selectActionRanges(actions, ctx);
    });
  }

  static async previewActions(actions: SheetAction[]): Promise<void> {
    try {
      const safeActions = guardActions(actions);
      if (!safeActions.length) return;

      await this.clearPreview();

      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const previews: PreviewRange[] = [];

        for (const action of safeActions) {
          try {
            const target = await this.getPreviewTarget(worksheet, action, context);
            if (!target) continue;

            target.range.load(['values', 'formulas', 'format/fill/color', 'format/fill/pattern']);
            await context.sync();

            previews.push({
              row: target.row,
              col: target.col,
              rowCount: target.rowCount,
              colCount: target.colCount,
              values: target.range.values,
              formulas: target.range.formulas,
              fillColor: target.range.format.fill.color ?? null,
              fillPattern: target.range.format.fill.pattern ?? 'None',
            });

            this.applyPreviewAction(target.range, action, target.colCount);
          } catch (error) {
            console.warn('[Cellix] Failed to preview action:', action, error);
          }
        }

        this.previewRanges = previews;
        await context.sync();
      });
    } catch (error) {
      console.warn('[Cellix] Failed to preview actions:', error);
    }
  }

  static async clearPreview(): Promise<void> {
    if (!this.previewRanges.length) return;

    const previews = [...this.previewRanges];
    this.previewRanges = [];

    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();

        for (const preview of previews) {
          const range = worksheet.getRangeByIndexes(
            preview.row,
            preview.col,
            preview.rowCount,
            preview.colCount,
          );
          range.values = preview.values as string[][];
          range.formulas = preview.formulas as string[][];
          this.restoreFill(range, preview.fillPattern, preview.fillColor);
        }

        await context.sync();
      });
    } catch (error) {
      console.warn('[Cellix] Failed to clear action preview:', error);
    }
  }

  static async commitPreview(): Promise<void> {
    if (!this.previewRanges.length) return;

    const previews = [...this.previewRanges];
    this.previewRanges = [];

    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();

        for (const preview of previews) {
          const range = worksheet.getRangeByIndexes(
            preview.row,
            preview.col,
            preview.rowCount,
            preview.colCount,
          );
          this.restoreFill(range, preview.fillPattern, preview.fillColor);
        }

        await context.sync();
      });
    } catch (error) {
      console.warn('[Cellix] Failed to commit action preview:', error);
    }
  }

  private static getRange(worksheet: Excel.Worksheet, action: SheetAction): Excel.Range {
    const row = action.row ?? 0;
    const col = action.col ?? 0;
    const rowCount = action.rowCount ?? 1;
    const colCount = action.colCount ?? 1;
    return worksheet.getRangeByIndexes(row, col, rowCount, colCount);
  }

  private static async getPreviewTarget(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<{ range: Excel.Range; row: number; col: number; rowCount: number; colCount: number } | null> {
    const cellTypes: SheetActionType[] = ['SET_CELL', 'CLEAR_CELL', 'SET_FORMULA', 'HIGHLIGHT_CELL'];
    if (cellTypes.includes(action.type) && action.row !== undefined && action.col !== undefined) {
      return {
        range: worksheet.getRangeByIndexes(action.row, action.col, 1, 1),
        row: action.row,
        col: action.col,
        rowCount: 1,
        colCount: 1,
      };
    }

    if (action.type === 'FORMAT_RANGE' && action.row !== undefined && action.col !== undefined) {
      return {
        range: this.getRange(worksheet, action),
        row: action.row,
        col: action.col,
        rowCount: action.rowCount ?? 1,
        colCount: action.colCount ?? 1,
      };
    }

    if (action.type === 'ADD_ROW') {
      const row = await this.getAppendRowIndex(worksheet, context);
      const colCount = Math.max(action.data?.length ?? 1, await this.getUsedColumnCount(worksheet, context));
      return { range: worksheet.getRangeByIndexes(row, 0, 1, colCount), row, col: 0, rowCount: 1, colCount };
    }

    if (action.type === 'DELETE_ROW' && action.row !== undefined) {
      const colCount = await this.getUsedColumnCount(worksheet, context);
      return {
        range: worksheet.getRangeByIndexes(action.row, 0, 1, colCount),
        row: action.row,
        col: 0,
        rowCount: 1,
        colCount,
      };
    }

    return null;
  }

  private static applyPreviewAction(range: Excel.Range, action: SheetAction, colCount: number): void {
    switch (action.type) {
      case 'SET_CELL':
        range.values = [[action.value ?? '']];
        return;
      case 'CLEAR_CELL':
        range.values = [['']];
        return;
      case 'SET_FORMULA':
        range.formulas = [[action.formula ?? '']];
        return;
      case 'ADD_ROW': {
        const data = action.data ?? [];
        range.values = [Array.from({ length: colCount }, (_, index) => data[index] ?? '')];
        return;
      }
      case 'HIGHLIGHT_CELL':
        if (action.color) this.applyFillColor(range, action.color);
        return;
      case 'FORMAT_RANGE':
        if (action.format?.fillColor) this.applyFillColor(range, action.format.fillColor);
        return;
      case 'DELETE_ROW':
        return;
      default:
        return;
    }
  }

  private static async getAppendRowIndex(
    worksheet: Excel.Worksheet,
    context: Excel.RequestContext,
  ): Promise<number> {
    const bounds = await this.getRealUsedBounds(worksheet, context);
    return bounds.nextRow;
  }

  private static async getUsedColumnCount(
    worksheet: Excel.Worksheet,
    context: Excel.RequestContext,
  ): Promise<number> {
    const bounds = await this.getRealUsedBounds(worksheet, context);
    return bounds.columnCount;
  }

  private static async getRealUsedBounds(
    worksheet: Excel.Worksheet,
    context: Excel.RequestContext,
  ): Promise<{ nextRow: number; columnCount: number }> {
    const usedRange = worksheet.getUsedRange();
    if (!usedRange) return { nextRow: 0, columnCount: 1 };

    usedRange.load(['values', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
    await context.sync();

    const values = usedRange.values ?? [];
    const baseRow = usedRange.rowIndex ?? 0;
    let lastRelativeRow = -1;
    let lastRelativeColumn = -1;

    values.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        if (!this.isBlankCell(cell)) {
          lastRelativeRow = Math.max(lastRelativeRow, rowIndex);
          lastRelativeColumn = Math.max(lastRelativeColumn, columnIndex);
        }
      });
    });

    const columnCount = Math.max(lastRelativeColumn + 1, 1);
    const firstRowHasContent = (values[0] ?? []).some((cell) => !this.isBlankCell(cell));

    if (lastRelativeRow < 0) {
      return { nextRow: firstRowHasContent ? baseRow + 1 : baseRow, columnCount };
    }

    let nextRow = baseRow + lastRelativeRow + 1;
    if (firstRowHasContent) {
      nextRow = Math.max(nextRow, baseRow + 1);
    }

    return { nextRow, columnCount };
  }

  private static isBlankCell(value: unknown): boolean {
    return value === null || value === undefined || String(value).trim() === '';
  }

  private static applyFillColor(range: Excel.Range, color: string): void {
    range.format.fill.pattern = 'Solid';
    range.format.fill.color = color;
  }

  private static restoreFill(
    range: Excel.Range,
    pattern: Excel.RangeFill['pattern'],
    color: string | null,
  ): void {
    range.format.fill.pattern = pattern;
    // A cell with no fill reports a null colour; clear() is how that is restored.
    if (color === null) {
      range.format.fill.clear();
      return;
    }
    range.format.fill.color = color;
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
