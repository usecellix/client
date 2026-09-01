import { SheetAction, SheetActionType } from '@/types/sheet-actions';
import { CellChange } from '@/types/changeSet';
import { ActionEngine } from '@/utils/actionEngine';
import type { CreatedConditionalFormatId, CreatedChartId } from '@/engine/actionEngine';
import {
  annotateDestOverwriteForCreatedSheets,
  pruneSpuriousAddSheets,
} from '@/engine/overwriteGuard';
import { sheetsCreatedInBatch } from '@/engine/sheetGuardState';

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

/**
 * Nothing in this class writes to the workbook before Accept — TASKS.md #148.
 *
 * It used to apply sheet-creation actions immediately on render, as a "soft
 * preview", so the diff view could live-read `before` values from sheets the
 * preview itself had just created. That single decision produced every Office.js
 * failure of the last several sessions (#145, #146, #147): code reading or
 * writing a sheet that existed only because an earlier, separate Excel call had
 * partially run, in an order nothing enforced.
 *
 * It was never necessary. For a sheet this batch creates, `before` is empty **by
 * construction** — `sheetsCreatedInBatch()` already proves exactly that, and the
 * diff can be computed rather than measured. Only edits to *pre-existing* sheets
 * need a live read, and that is a read, never a write.
 *
 * Consequences, all of them good:
 *   - Reject is a genuine no-op. Nothing happened, so nothing is undone.
 *     `buildPreviewRejectActions`'s revert machinery is no longer reachable from
 *     here (kept for `ChangeHistoryPanel`/`CheckpointPanel`, which revert real,
 *     applied change sets).
 *   - Accept applies the complete batch in one pass, in plan order, so a build
 *     can never be left half-created (the #141 failure shape).
 *   - ARCHITECTURE.md AD-1 is unchanged: `RichActionEngine` is still the only
 *     writer. There is simply now exactly one moment at which it runs.
 */
export class PreviewManager {
  private pendingActions: SheetAction[] = [];
  private isActive = false;
  private applying = false;

  async render(payload: ActionsPayload): Promise<DiffItem[]> {
    if (this.isActive) {
      await this.reject();
    }

    this.pendingActions = payload.actions;
    this.isActive = true;
    const diffItems: DiffItem[] = [];

    const createdHere = sheetsCreatedInBatch(payload.actions);

    await Excel.run(async (ctx) => {
      const activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      await ctx.sync();

      for (const action of payload.actions) {
        // COPY/sort natives use sourceSheet/destSheet — skip if no cell target.
        const sheetName =
          action.sheetName ??
          (action as { destSheet?: string }).destSheet ??
          activeSheet.name;
        if (!sheetName) continue;

        // A sheet this batch creates is empty by construction, so its `before`
        // is knowable without asking Excel — and asking would be worse than
        // pointless, since the sheet does not exist yet and every accessor that
        // reaches for it throws. Compute, don't measure. TASKS.md #148.
        if (createdHere.has(String(sheetName).trim().toLowerCase())) {
          const predicted = this.predictedTarget(action);
          if (predicted) {
            diffItems.push({
              sheetName: String(sheetName),
              address: predicted.address,
              actionType: action.type,
              before: '(empty)',
              after: this.describeAfter(action),
              description: this.autoDescription(action, predicted.address),
            });
          }
          continue;
        }

        const ws = ctx.workbook.worksheets.getItemOrNullObject(String(sheetName));
        ws.load('isNullObject,name');
        await ctx.sync();
        if (ws.isNullObject) {
          // Not created by this batch and not present: nothing to diff against.
          continue;
        }

        const target = await this.resolveTarget(ws, action, ctx);
        if (!target) continue;

        const { range, row, col, rowCount, colCount } = target;
        const resolvedAddress = formatAddress(row, col, rowCount, colCount);

        range.load(['values']);
        await ctx.sync();

        diffItems.push({
          sheetName: String(sheetName),
          address: resolvedAddress,
          actionType: action.type,
          before: this.serializeValues(range.values),
          after: this.describeAfter(action),
          description: this.autoDescription(action, resolvedAddress),
        });
      }

      await ctx.sync();
    });

    return diffItems;
  }

  async accept(): Promise<
    | {
        createdConditionalFormatIds?: CreatedConditionalFormatId[];
        createdChartIds?: CreatedChartId[];
        sortedRangeChanges?: CellChange[];
      }
    | void
  > {
    if (!this.isActive || this.applying) return;
    this.applying = true;

    // The whole batch, in plan order — nothing was applied early, so there is
    // nothing to subtract. `splitIntoActionWaves` already hoisted sheet creates
    // to the front server-side, so creates still precede the writes that need
    // them. The old `structuralApplied`/`deferredApplied` bookkeeping existed
    // only to avoid re-running what the soft preview had already done; with no
    // soft preview there is nothing to double-apply. TASKS.md #148.
    const toApply = annotateDestOverwriteForCreatedSheets(
      pruneSpuriousAddSheets([...this.pendingActions]),
    );

    try {
      let createdConditionalFormatIds: CreatedConditionalFormatId[] | undefined;
      let createdChartIds: CreatedChartId[] | undefined;
      let sortedRangeChanges: CellChange[] | undefined;
      if (toApply.length > 0) {
        // applyActionsWithReport (not the void applyActions) — TASKS.md #40/#15 need
        // their createdConditionalFormatIds/createdChartIds. Same "errors present +
        // nothing applied = throw" behavior as applyActions, replicated here so
        // Accept's error handling is unchanged.
        const result = await ActionEngine.applyActionsWithReport(toApply);
        if (result.errors.length > 0 && result.applied === 0) {
          throw new Error(result.errors.join('; '));
        }
        createdConditionalFormatIds = result.createdConditionalFormatIds;
        createdChartIds = result.createdChartIds;
        sortedRangeChanges = result.sortedRangeChanges;
      }
      this.reset();
      return createdConditionalFormatIds || createdChartIds || sortedRangeChanges
        ? { createdConditionalFormatIds, createdChartIds, sortedRangeChanges }
        : undefined;
    } catch (error) {
      // Keep pending + applied flags so a retry does not double-write.
      throw error;
    } finally {
      this.applying = false;
    }
  }

  /**
   * Reject is now a genuine no-op: no write ever happened before Accept, so
   * there is nothing to undo. This is the whole point of TASKS.md #148 — the
   * previous implementation had to reverse its own soft preview, and a revert
   * that itself fails (or partially succeeds) is a strictly worse failure mode
   * than never having written.
   */
  async reject(): Promise<void> {
    if (!this.isActive || this.applying) return;
    this.reset();
  }

  // `changes` is accepted for call-site compatibility but no longer stored:
  // nothing is highlighted before Accept, so there is no pending-highlight state
  // to keep. TASKS.md #148.
  async highlightChanges(_changes: CellChange[], actions: SheetAction[] = []): Promise<void> {
    if (this.isActive) {
      await this.reject();
    }

    this.pendingActions = actions;
    this.isActive = true;
  }

  get active(): boolean {
    return this.isActive;
  }

  private reset(): void {
    this.pendingActions = [];
    this.isActive = false;
  }

  /**
   * Where an action will write, derived from the action alone.
   *
   * Used for sheets this batch creates, where no live read is possible (the
   * sheet does not exist yet) and none is needed (it will be empty). Mirrors
   * `resolveTarget`'s addressing, minus everything that requires a live range.
   * `ADD_ROW`/`APPEND_ROW` land at row 0 on a sheet created moments earlier —
   * that is exactly what `handleAppendRow` computes at apply time against a
   * used range that does not yet exist. TASKS.md #148.
   */
  private predictedTarget(action: SheetAction): { address: string } | null {
    if (
      CELL_ACTION_TYPES.includes(action.type) &&
      action.row !== undefined &&
      action.col !== undefined
    ) {
      return {
        address: formatAddress(
          action.row,
          action.col,
          action.rowCount ?? 1,
          action.colCount ?? 1,
        ),
      };
    }

    if (action.type === 'ADD_ROW' || action.type === 'APPEND_ROW') {
      const values = (action.data ?? action.values) as unknown[] | undefined;
      const colCount = Math.max(Array.isArray(values) ? values.length : 1, 1);
      return { address: formatAddress(0, 0, 1, colCount) };
    }

    if (action.type === 'WRITE_TABLE') {
      const headers = action.headers ?? [];
      const rows = action.rows ?? [];
      const rowCount = headers.length ? rows.length + 1 : rows.length;
      if (!rowCount) return null;
      const colCount = Math.max(
        headers.length,
        ...rows.map((row) => (Array.isArray(row) ? row.length : 0)),
        1,
      );
      return { address: formatAddress(0, 0, rowCount, colCount) };
    }

    return null;
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
    // `getUsedRange()` throws ItemNotFound on a sheet with nothing on it —
    // the exact state of a sheet a preceding ADD_SHEET in the same batch just
    // created. This is reached whenever the preview's diff loop (below)
    // resolves an ADD_ROW target, which an ADD_ROW-shaped header write does
    // on the very sheet it's about to populate — a real, reproduced live
    // failure: "The requested resource doesn't exist." before any sheet got
    // its content, or (once staged accept waves were merged into one in
    // TASKS.md #141) before Accept did anything at all, since this runs at
    // preview time. `getUsedRangeOrNullObject` is the null-safe sibling.
    // TASKS.md #146.
    const usedRange = worksheet.getUsedRangeOrNullObject();
    usedRange.load(['values', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount', 'isNullObject']);
    await context.sync();
    if (usedRange.isNullObject) return { nextRow: 0, columnCount: 1 };

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
