import { SheetAction, SheetActionType } from '@/types/sheet-actions';
import { CellChange } from '@/types/changeSet';
import { ActionEngine } from '@/utils/actionEngine';
import {
  buildPreviewRejectActions,
  partitionPreviewActions,
} from '@/utils/previewRevert';

/* global Excel */

export interface ActionsPayload {
  actions: SheetAction[];
  summary: string;
  affectedCells?: string[];
}

export interface DiffItem {
  sheetName: string;
  address: string;
  actionType: string;
  before: string;
  after: string;
  description: string;
}

const CELL_ACTION_TYPES: SheetActionType[] = [
  'SET_CELL',
  'CLEAR_CELL',
  'SET_FORMULA',
  'HIGHLIGHT_CELL',
  'FORMAT_RANGE',
];

export class PreviewManager {
  private pendingActions: SheetAction[] = [];
  private pendingChanges: CellChange[] = [];
  private structuralApplied = false;
  private deferredApplied = false;
  private isActive = false;
  private applying = false;

  async render(payload: ActionsPayload): Promise<DiffItem[]> {
    if (this.isActive) {
      await this.reject();
    }

    this.pendingActions = payload.actions;
    this.pendingChanges = [];
    this.structuralApplied = false;
    this.deferredApplied = false;
    const diffItems: DiffItem[] = [];

    await this.applyStructuralPreview(payload.actions);

    await Excel.run(async (ctx) => {
      const activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      await ctx.sync();

      for (const action of payload.actions) {
        const sheetName = action.sheetName ?? activeSheet.name;
        const ws = ctx.workbook.worksheets.getItem(sheetName);
        const target = await this.resolveTarget(ws, action, ctx);
        if (!target) continue;

        const { range, row, col, rowCount, colCount } = target;
        const resolvedAddress = formatAddress(row, col, rowCount, colCount);

        range.load(['values']);
        await ctx.sync();

        diffItems.push({
          sheetName,
          address: resolvedAddress,
          actionType: action.type,
          before: this.serializeValues(range.values),
          after: this.describeAfter(action),
          description: this.autoDescription(action, resolvedAddress),
        });
      }

      await ctx.sync();
    });

    this.isActive = true;
    return diffItems;
  }

  async accept(): Promise<void> {
    if (!this.isActive || this.applying) return;
    this.applying = true;

    const actions = [...this.pendingActions];
    const { deferred } = partitionPreviewActions(actions);

    this.pendingActions = [];
    this.pendingChanges = [];
    this.isActive = false;

    try {
      if (deferred.length > 0) {
        await ActionEngine.applyActions(deferred);
        this.deferredApplied = true;
      }
    } catch (error) {
      this.pendingActions = actions;
      this.isActive = true;
      throw error;
    } finally {
      this.structuralApplied = false;
      this.deferredApplied = false;
      this.applying = false;
    }
  }

  async reject(): Promise<void> {
    if (!this.isActive || this.applying) return;
    this.applying = true;

    const actions = [...this.pendingActions];
    const changes = [...this.pendingChanges];
    const hadStructural = this.structuralApplied;
    const hadDeferred = this.deferredApplied;

    try {
      const revertActions = buildPreviewRejectActions(actions, changes, {
        structuralApplied: hadStructural,
        deferredApplied: hadDeferred,
      });

      if (revertActions.length > 0) {
        await ActionEngine.applyActions(revertActions);
      }
    } finally {
      this.reset();
      this.applying = false;
    }
  }

  async highlightChanges(changes: CellChange[], actions: SheetAction[] = []): Promise<void> {
    if (this.isActive) {
      await this.reject();
    }

    this.pendingActions = actions;
    this.pendingChanges = changes;
    this.structuralApplied = false;
    this.deferredApplied = false;

    await this.applyStructuralPreview(actions);
    this.isActive = true;
  }

  get active(): boolean {
    return this.isActive;
  }

  private reset(): void {
    this.pendingActions = [];
    this.pendingChanges = [];
    this.structuralApplied = false;
    this.deferredApplied = false;
    this.isActive = false;
  }

  /** Apply sheet/structural mutations immediately so reject can undo them. */
  private async applyStructuralPreview(actions: SheetAction[]): Promise<void> {
    const { structural } = partitionPreviewActions(actions);
    if (structural.length === 0) return;

    await ActionEngine.applyActions(structural);
    this.structuralApplied = true;
  }

  private async resolveTarget(
    worksheet: Excel.Worksheet,
    action: SheetAction,
    context: Excel.RequestContext,
  ): Promise<{
    range: Excel.Range;
    row: number;
    col: number;
    rowCount: number;
    colCount: number;
  } | null> {
    if (
      CELL_ACTION_TYPES.includes(action.type) &&
      action.row !== undefined &&
      action.col !== undefined
    ) {
      const rowCount = action.rowCount ?? 1;
      const colCount = action.colCount ?? 1;
      return {
        range: worksheet.getRangeByIndexes(action.row, action.col, rowCount, colCount),
        row: action.row,
        col: action.col,
        rowCount,
        colCount,
      };
    }

    if (action.type === 'ADD_ROW') {
      const row = await this.getAppendRowIndex(worksheet, context);
      const colCount = Math.max(
        action.data?.length ?? 1,
        await this.getUsedColumnCount(worksheet, context),
      );
      return {
        range: worksheet.getRangeByIndexes(row, 0, 1, colCount),
        row,
        col: 0,
        rowCount: 1,
        colCount,
      };
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

    if (action.type === 'INSERT_ROW' && action.row !== undefined) {
      const colCount = await this.getUsedColumnCount(worksheet, context);
      const count = action.count ?? 1;
      return {
        range: worksheet.getRangeByIndexes(action.row, 0, count, colCount),
        row: action.row,
        col: 0,
        rowCount: count,
        colCount,
      };
    }

    if (action.type === 'WRITE_TABLE') {
      const headers = action.headers ?? [];
      const rows = action.rows ?? [];
      const rowCount = headers.length ? rows.length + 1 : 0;
      const colCount = Math.max(
        headers.length,
        ...rows.map((row) => (Array.isArray(row) ? row.length : 0)),
        1,
      );
      if (!rowCount) return null;
      return {
        range: worksheet.getRangeByIndexes(0, 0, rowCount, colCount),
        row: 0,
        col: 0,
        rowCount,
        colCount,
      };
    }

    return null;
  }

  private async getAppendRowIndex(
    worksheet: Excel.Worksheet,
    context: Excel.RequestContext,
  ): Promise<number> {
    const bounds = await this.getRealUsedBounds(worksheet, context);
    return bounds.nextRow;
  }

  private async getUsedColumnCount(
    worksheet: Excel.Worksheet,
    context: Excel.RequestContext,
  ): Promise<number> {
    const bounds = await this.getRealUsedBounds(worksheet, context);
    return bounds.columnCount;
  }

  private async getRealUsedBounds(
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

  private isBlankCell(value: unknown): boolean {
    return value === null || value === undefined || String(value).trim() === '';
  }

  private serializeValues(values: unknown[][] | undefined): string {
    if (!values || values.length === 0) return '(empty)';
    return values.flat().slice(0, 5).join(', ');
  }

  private describeAfter(action: SheetAction): string {
    if (action.value != null) return String(action.value);
    if (action.data) return `[${action.data.length} cells]`;
    if (action.formula) return action.formula;
    if (action.headers) return `[table: ${action.headers.join(', ')}]`;
    return action.type;
  }

  private autoDescription(action: SheetAction, address: string): string {
    switch (action.type) {
      case 'SET_CELL':
        return `Set ${address} to "${action.value ?? ''}"`;
      case 'CLEAR_CELL':
        return `Clear ${address}`;
      case 'ADD_ROW':
        return `Add row at ${address}`;
      case 'DELETE_ROW':
        return `Delete row ${(action.row ?? 0) + 1}`;
      case 'SET_FORMULA':
        return `Set formula ${action.formula} at ${address}`;
      case 'FORMAT_RANGE':
        return `Format ${address}`;
      case 'WRITE_TABLE':
        return `Write table at ${address}`;
      default:
        return action.type;
    }
  }
}

function colIndexToLetter(col: number): string {
  let index = col + 1;
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

function formatAddress(row: number, col: number, rowCount: number, colCount: number): string {
  const start = `${colIndexToLetter(col)}${row + 1}`;
  if (rowCount === 1 && colCount === 1) return start;
  const end = `${colIndexToLetter(col + colCount - 1)}${row + rowCount}`;
  return `${start}:${end}`;
}

export const previewManager = new PreviewManager();
