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
  fillPattern: Excel.FillPattern;
};

const PREVIEW_FILL = '#DCFCE7';

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
    try {
      const safeActions = guardActions(actions);
      if (!safeActions.length) {
        throw new Error(
          'No safe actions to apply. Header row is protected — use ADD_ROW to append data.',
        );
      }

      const actionsToApply = this.previewRanges.length
        ? safeActions.filter((action) => action.type === 'DELETE_ROW' || action.type === 'HIGHLIGHT_CELL')
        : safeActions;

      if (this.previewRanges.length) {
        await this.commitPreview();
      }

      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();

        for (const action of actionsToApply) {
          try {
            await this.applySingleAction(context, worksheet, action);
            if (action.type === 'ADD_ROW') {
              await context.sync();
            }
          } catch (error) {
            console.warn('Action failed, skipping:', action, error);
          }
        }

        await context.sync();
      });
    } catch (error) {
      console.error('Failed to apply actions to spreadsheet:', error);
      throw new Error(`Spreadsheet update failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
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

  private static async applySingleAction(
    context: Excel.RequestContext,
    worksheet: Excel.Worksheet,
    action: SheetAction,
  ): Promise<void> {
    switch (action.type) {
      case 'SET_CELL':
        this.setCell(worksheet, action);
        break;
      case 'CLEAR_CELL':
        this.clearCell(worksheet, action);
        break;
      case 'HIGHLIGHT_CELL':
        this.highlightCell(worksheet, action);
        break;
      case 'SET_FORMULA':
        this.setFormula(worksheet, action);
        break;
      case 'ADD_ROW':
        await this.addRow(worksheet, action, context);
        break;
      case 'DELETE_ROW':
        await this.deleteRow(worksheet, action, context);
        break;
      case 'INSERT_ROW':
        await this.insertRow(worksheet, action, context);
        break;
      case 'INSERT_COLUMN':
        await this.insertColumn(worksheet, action, context);
        break;
      case 'DELETE_COLUMN':
        await this.deleteColumn(worksheet, action, context);
        break;
      case 'HIDE_ROW':
        this.hideRow(worksheet, action);
        break;
      case 'SHOW_ROW':
        this.showRow(worksheet, action);
        break;
      case 'HIDE_COLUMN':
        this.hideColumn(worksheet, action);
        break;
      case 'SHOW_COLUMN':
        this.showColumn(worksheet, action);
        break;
      case 'SET_ROW_HEIGHT':
        this.setRowHeight(worksheet, action);
        break;
      case 'SET_COLUMN_WIDTH':
        this.setColumnWidth(worksheet, action);
        break;
      case 'FREEZE_PANES':
        this.freezePanes(worksheet, action);
        break;
      case 'UNFREEZE_PANES':
        worksheet.freezePanes.unfreeze();
        break;
      case 'MERGE_CELLS':
        this.mergeCells(worksheet, action);
        break;
      case 'UNMERGE_CELLS':
        this.unmergeCells(worksheet, action);
        break;
      case 'CLEAR_CONTENT':
        this.clearRange(worksheet, action, Excel.ClearApplyTo.contents);
        break;
      case 'CLEAR_FORMAT':
        this.clearRange(worksheet, action, Excel.ClearApplyTo.formats);
        break;
      case 'CLEAR_ALL':
        this.clearRange(worksheet, action, Excel.ClearApplyTo.all);
        break;
      case 'FORMAT_RANGE':
        this.formatRange(worksheet, action);
        break;
      case 'FILL_DOWN':
        await this.fillDown(worksheet, action, context);
        break;
      case 'FILL_RIGHT':
        await this.fillRight(worksheet, action, context);
        break;
      case 'CREATE_SHEET':
        await this.createSheet(context, action);
        break;
      case 'DELETE_SHEET':
        await this.deleteSheet(context, action);
        break;
      case 'RENAME_SHEET':
        worksheet.name = action.newSheetName ?? action.sheetName ?? worksheet.name;
        break;
      case 'COPY_SHEET':
        await this.copySheet(context, action);
        break;
      case 'HIDE_SHEET':
        await this.setSheetVisibility(context, action.sheetName, Excel.SheetVisibility.hidden);
        break;
      case 'SHOW_SHEET':
        await this.setSheetVisibility(context, action.sheetName, Excel.SheetVisibility.visible);
        break;
      case 'SET_SHEET_COLOR':
        await this.setSheetColor(context, action);
        break;
      case 'ADD_COMMENT':
        this.addComment(worksheet, action);
        break;
      case 'DELETE_COMMENT':
        this.deleteComment(worksheet, action);
        break;
      case 'WRITE_TABLE':
        this.writeTable(worksheet, action);
        break;
      default:
        console.warn(`Unknown action type: ${(action as SheetAction).type}`);
    }
  }

  private static getRange(worksheet: Excel.Worksheet, action: SheetAction): Excel.Range {
    const row = action.row ?? 0;
    const col = action.col ?? 0;
    const rowCount = action.rowCount ?? 1;
    const colCount = action.colCount ?? 1;
    return worksheet.getRangeByIndexes(row, col, rowCount, colCount);
  }

  private static setCell(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined || action.col === undefined || action.value === undefined) {
      throw new Error('SET_CELL action requires row, col, and value');
    }
    this.getRange(worksheet, { ...action, rowCount: 1, colCount: 1 }).values = [[action.value]];
  }

  private static clearCell(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined || action.col === undefined) {
      throw new Error('CLEAR_CELL action requires row and col');
    }
    this.getRange(worksheet, { ...action, rowCount: 1, colCount: 1 }).values = [['']];
  }

  private static highlightCell(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined || action.col === undefined) {
      throw new Error('HIGHLIGHT_CELL action requires row and col');
    }
    this.applyFillColor(
      this.getRange(worksheet, { ...action, rowCount: 1, colCount: 1 }),
      action.color || PREVIEW_FILL,
    );
  }

  private static setFormula(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined || action.col === undefined || action.formula === undefined) {
      throw new Error('SET_FORMULA action requires row, col, and formula');
    }
    this.getRange(worksheet, { ...action, rowCount: 1, colCount: 1 }).formulas = [[action.formula]];
  }

  private static writeTable(worksheet: Excel.Worksheet, action: SheetAction): void {
    const headers = action.headers ?? [];
    const rows = action.rows ?? [];
    if (!headers.length) {
      throw new Error('WRITE_TABLE requires headers');
    }

    const tableRows = [headers, ...rows];
    const rowCount = tableRows.length;
    const colCount = Math.max(
      headers.length,
      ...rows.map((row) => (Array.isArray(row) ? row.length : 0)),
      1,
    );

    const range = worksheet.getRangeByIndexes(0, 0, rowCount, colCount);
    range.values = tableRows.map((row) =>
      Array.from({ length: colCount }, (_, index) =>
        Array.isArray(row) ? (row[index] ?? '') : '',
      ),
    );
  }

  private static async addRow(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<void> {
    if (!Array.isArray(action.data)) {
      throw new Error('ADD_ROW action requires data array');
    }

    const targetRow = await this.getAppendRowIndex(worksheet, context);
    const colCount = Math.max(action.data.length, await this.getUsedColumnCount(worksheet, context));
    const rowRange = worksheet.getRangeByIndexes(targetRow, 0, 1, colCount);
    rowRange.getEntireRow().insert('Down');
    rowRange.values = [Array.from({ length: colCount }, (_, index) => action.data![index] ?? '')];
  }

  private static async insertRow(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<void> {
    const row = action.row ?? (await this.getAppendRowIndex(worksheet, context));
    const count = action.count ?? 1;
    const colCount = await this.getUsedColumnCount(worksheet, context);
    const range = worksheet.getRangeByIndexes(row, 0, count, colCount);
    range.getEntireRow().insert(action.position === 'above' ? 'Up' : 'Down');
  }

  private static async deleteRow(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<void> {
    if (action.row === undefined) throw new Error('DELETE_ROW action requires row');
    const columnCount = await this.getUsedColumnCount(worksheet, context);
    worksheet.getRangeByIndexes(action.row, 0, 1, columnCount).delete('Up');
  }

  private static async insertColumn(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<void> {
    if (action.col === undefined) throw new Error('INSERT_COLUMN action requires col');
    const count = action.count ?? 1;
    const rowCount = await this.getUsedRowCount(worksheet, context);
    const range = worksheet.getRangeByIndexes(0, action.col, rowCount, count);
    range.getEntireColumn().insert(action.position === 'left' ? 'Left' : 'Right');
  }

  private static async deleteColumn(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<void> {
    if (action.col === undefined) throw new Error('DELETE_COLUMN action requires col');
    const rowCount = await this.getUsedRowCount(worksheet, context);
    worksheet.getRangeByIndexes(0, action.col, rowCount, 1).delete('Left');
  }

  private static hideRow(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined) return;
    worksheet.getRangeByIndexes(action.row, 0, action.rowCount ?? 1, 1).rowHidden = true;
  }

  private static showRow(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined) return;
    worksheet.getRangeByIndexes(action.row, 0, action.rowCount ?? 1, 1).rowHidden = false;
  }

  private static hideColumn(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.col === undefined) return;
    worksheet.getRangeByIndexes(0, action.col, 1, action.colCount ?? 1).columnHidden = true;
  }

  private static showColumn(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.col === undefined) return;
    worksheet.getRangeByIndexes(0, action.col, 1, action.colCount ?? 1).columnHidden = false;
  }

  private static setRowHeight(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined || action.height === undefined) return;
    worksheet.getRangeByIndexes(action.row, 0, 1, 1).format.rowHeight = action.height;
  }

  private static setColumnWidth(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.col === undefined || action.width === undefined) return;
    worksheet.getRangeByIndexes(0, action.col, 1, 1).format.columnWidth = action.width;
  }

  private static freezePanes(worksheet: Excel.Worksheet, action: SheetAction): void {
    const rows = action.freezeRows ?? 1;
    const cols = action.freezeColumns ?? 0;
    if (rows > 0) worksheet.freezePanes.freezeRows(rows);
    if (cols > 0) worksheet.freezePanes.freezeColumns(cols);
  }

  private static mergeCells(worksheet: Excel.Worksheet, action: SheetAction): void {
    this.getRange(worksheet, action).merge(action.mergeAcross ?? false);
  }

  private static unmergeCells(worksheet: Excel.Worksheet, action: SheetAction): void {
    this.getRange(worksheet, action).unmerge();
  }

  private static clearRange(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    applyTo: Excel.ClearApplyTo,
  ): void {
    this.getRange(worksheet, action).clear(applyTo);
  }

  private static formatRange(worksheet: Excel.Worksheet, action: SheetAction): void {
    const range = this.getRange(worksheet, action);
    const fmt = action.format;
    if (!fmt) return;

    if (fmt.bold !== undefined) range.format.font.bold = fmt.bold;
    if (fmt.italic !== undefined) range.format.font.italic = fmt.italic;
    if (fmt.underline !== undefined) range.format.font.underline = fmt.underline ? 'Single' : 'None';
    if (fmt.fontSize !== undefined) range.format.font.size = fmt.fontSize;
    if (fmt.fontColor !== undefined) range.format.font.color = fmt.fontColor;
    if (fmt.fillColor !== undefined) {
      range.format.fill.pattern = 'Solid';
      range.format.fill.color = fmt.fillColor;
    }
    if (fmt.horizontalAlignment !== undefined) {
      const map: Record<string, Excel.HorizontalAlignment> = {
        left: 'Left',
        center: 'Center',
        right: 'Right',
      };
      range.format.horizontalAlignment = map[fmt.horizontalAlignment] ?? 'General';
    }
    if (fmt.verticalAlignment !== undefined) {
      const map: Record<string, Excel.VerticalAlignment> = {
        top: 'Top',
        middle: 'Center',
        bottom: 'Bottom',
      };
      range.format.verticalAlignment = map[fmt.verticalAlignment] ?? 'Bottom';
    }
    if (fmt.wrapText !== undefined) range.format.wrapText = fmt.wrapText;
    if (fmt.numberFormat !== undefined) range.numberFormat = [[fmt.numberFormat]];

    if (fmt.borders) {
      const borders = range.format.borders;
      const style = 'Continuous' as Excel.BorderStyle;
      if (fmt.borders === 'all' || fmt.borders === 'outer') {
        (['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] as Excel.BorderSide[]).forEach((edge) => {
          borders.getItem(edge).style = style;
        });
      }
      if (fmt.borders === 'bottom') {
        borders.getItem('EdgeBottom').style = style;
      }
      if (fmt.borders === 'none') {
        (['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] as Excel.BorderSide[]).forEach((edge) => {
          borders.getItem(edge).style = 'None';
        });
      }
    }
  }

  private static async fillDown(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<void> {
    if (action.col === undefined || action.row === undefined) return;
    const endRow = action.endRow ?? (await this.getUsedRowCount(worksheet, context)) - 1;
    if (endRow <= action.row) return;

    const source = worksheet.getRangeByIndexes(action.row, action.col, 1, 1);
    const dest = worksheet.getRangeByIndexes(action.row, action.col, endRow - action.row + 1, 1);
    source.autoFill(dest, Excel.AutoFillType.fillDefault);
  }

  private static async fillRight(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<void> {
    if (action.row === undefined || action.col === undefined) return;
    const endCol = action.endCol ?? (await this.getUsedColumnCount(worksheet, context)) - 1;
    if (endCol <= action.col) return;

    const source = worksheet.getRangeByIndexes(action.row, action.col, 1, 1);
    const dest = worksheet.getRangeByIndexes(action.row, action.col, 1, endCol - action.col + 1);
    source.autoFill(dest, Excel.AutoFillType.fillDefault);
  }

  private static async createSheet(context: Excel.RequestContext, action: SheetAction): Promise<void> {
    const sheets = context.workbook.worksheets;
    const newSheet = sheets.add(action.sheetName ?? 'New Sheet');

    if (action.relativeTo && action.position) {
      const refSheet = sheets.getItem(action.relativeTo);
      refSheet.load('position');
      await context.sync();
      newSheet.position = action.position === 'before' ? refSheet.position : refSheet.position + 1;
    }
  }

  private static async deleteSheet(context: Excel.RequestContext, action: SheetAction): Promise<void> {
    if (!action.sheetName) return;
    context.workbook.worksheets.getItem(action.sheetName).delete();
  }

  private static async copySheet(context: Excel.RequestContext, action: SheetAction): Promise<void> {
    if (!action.sheetName) return;
    const copy = context.workbook.worksheets.getItem(action.sheetName).copy();
    if (action.newSheetName) copy.name = action.newSheetName;
  }

  private static async setSheetVisibility(
    context: Excel.RequestContext,
    sheetName: string | undefined,
    visibility: Excel.SheetVisibility,
  ): Promise<void> {
    if (!sheetName) return;
    context.workbook.worksheets.getItem(sheetName).visibility = visibility;
  }

  private static async setSheetColor(context: Excel.RequestContext, action: SheetAction): Promise<void> {
    if (!action.sheetName || !action.color) return;
    context.workbook.worksheets.getItem(action.sheetName).tabColor = action.color;
  }

  private static addComment(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined || action.col === undefined || !action.comment) return;
    this.getRange(worksheet, { ...action, rowCount: 1, colCount: 1 }).getComment().add(action.comment);
  }

  private static deleteComment(worksheet: Excel.Worksheet, action: SheetAction): void {
    if (action.row === undefined || action.col === undefined) return;
    this.getRange(worksheet, { ...action, rowCount: 1, colCount: 1 }).getComment().delete();
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
        this.applyFillColor(range, PREVIEW_FILL);
        return;
      case 'CLEAR_CELL':
        range.values = [['']];
        this.applyFillColor(range, PREVIEW_FILL);
        return;
      case 'SET_FORMULA':
        range.formulas = [[action.formula ?? '']];
        this.applyFillColor(range, PREVIEW_FILL);
        return;
      case 'ADD_ROW': {
        const data = action.data ?? [];
        range.values = [Array.from({ length: colCount }, (_, index) => data[index] ?? '')];
        this.applyFillColor(range, PREVIEW_FILL);
        return;
      }
      case 'HIGHLIGHT_CELL':
        this.applyFillColor(range, action.color || PREVIEW_FILL);
        return;
      case 'FORMAT_RANGE':
        if (action.format?.fillColor) this.applyFillColor(range, action.format.fillColor);
        else this.applyFillColor(range, PREVIEW_FILL);
        return;
      case 'DELETE_ROW':
        this.applyFillColor(range, PREVIEW_FILL);
        return;
      default:
        this.applyFillColor(range, PREVIEW_FILL);
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

  private static async getUsedRowCount(
    worksheet: Excel.Worksheet,
    context: Excel.RequestContext,
  ): Promise<number> {
    const bounds = await this.getRealUsedBounds(worksheet, context);
    return bounds.nextRow;
  }

  private static async getRealUsedBounds(
    worksheet: Excel.Worksheet,
    context: Excel.RequestContext,
  ): Promise<{ nextRow: number; columnCount: number }> {
    const usedRange = worksheet.getUsedRange();
    if (!usedRange) return { nextRow: 0, columnCount: 1 };

    usedRange.load(['values', 'row', 'column', 'rowCount', 'columnCount']);
    await context.sync();

    const values = usedRange.values ?? [];
    const baseRow = usedRange.row ?? 0;
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
    pattern: Excel.FillPattern,
    color: string | null,
  ): void {
    range.format.fill.pattern = pattern;
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
      case 'SHOW_ROW':
      case 'SET_ROW_HEIGHT':
        return action.row !== undefined;
      case 'INSERT_COLUMN':
      case 'DELETE_COLUMN':
      case 'HIDE_COLUMN':
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
      default:
        return false;
    }
  }
}

export type { SheetAction };
