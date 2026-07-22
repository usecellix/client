import { RichAction } from '@/action.types';
import { columnIndexToLetter, parseCellAddress } from './addressUtils';
import { resolveWorksheet } from './handlers/resolveWorksheet';

/* global Excel */

export class OverwriteGuardError extends Error {
  readonly targetRange: string;
  readonly sampleExistingValues: unknown[][];

  constructor(opts: {
    message: string;
    targetRange: string;
    sampleExistingValues: unknown[][];
  }) {
    super(opts.message);
    this.name = 'OverwriteGuardError';
    this.targetRange = opts.targetRange;
    this.sampleExistingValues = opts.sampleExistingValues;
  }
}

export function isOverwriteGuardError(error: unknown): error is OverwriteGuardError {
  return error instanceof OverwriteGuardError ||
    (Boolean(error) &&
      typeof error === 'object' &&
      (error as { name?: string }).name === 'OverwriteGuardError');
}

/** True when any cell has a non-empty value (string/number/boolean). */
export function rangeHasExistingData(values: unknown[][] | undefined | null): boolean {
  if (!values?.length) return false;
  return values.some((row) =>
    Array.isArray(row) && row.some((cell) => cell !== '' && cell != null),
  );
}

export function hasExplicitOverwriteConfirmed(action: RichAction): boolean {
  return Boolean((action as { explicitOverwriteConfirmed?: boolean }).explicitOverwriteConfirmed);
}

/**
 * Write actions that can destroy existing cell values.
 * Format-only / structural-shift / intentional-clear actions are excluded.
 */
export const OVERWRITE_GUARDED_ACTION_TYPES = new Set<RichAction['type']>([
  'SET_CELL',
  'SET_FORMULA',
  'BATCH_SET',
  'WRITE_TABLE',
  'FILL_DOWN',
  'AUTO_FILL',
  'APPEND_ROW',
  'COPY_FILTERED_RANGE',
  'MOVE_RANGE',
  'AGGREGATE_TABLE',
]);

export function isOverwriteGuardedAction(action: RichAction): boolean {
  return OVERWRITE_GUARDED_ACTION_TYPES.has(action.type);
}

export type ResolvedWriteTarget = {
  sheetName: string;
  address: string;
  /** Optional pre-resolved Excel range; when set, guard loads values from it. */
  getRange: (sheet: Excel.Worksheet) => Excel.Range;
};

function singleCellAddress(row: number, col: number): string {
  return `${columnIndexToLetter(col)}${row + 1}`;
}

function rangeAddressFromIndexes(
  row: number,
  col: number,
  rowCount: number,
  colCount: number,
): string {
  const start = singleCellAddress(row, col);
  if (rowCount <= 1 && colCount <= 1) return start;
  return `${start}:${singleCellAddress(row + rowCount - 1, col + colCount - 1)}`;
}

/**
 * Resolve the cell/range this action will write values into.
 * Returns null when the action does not write values (or target cannot be known yet).
 */
export function resolveWriteTarget(action: RichAction): ResolvedWriteTarget | null {
  switch (action.type) {
    case 'SET_CELL':
    case 'SET_FORMULA':
      return {
        sheetName: action.sheetName,
        address: action.address,
        getRange: (sheet) => sheet.getRange(action.address),
      };
    case 'BATCH_SET': {
      if (!action.operations?.length) return null;
      const addresses = action.operations.map((op) => op.address).filter(Boolean);
      if (!addresses.length) return null;
      // Guard each op independently via a synthetic multi-check in guardAgainstOverwrite.
      return {
        sheetName: action.sheetName,
        address: addresses.join(','),
        getRange: (sheet) => sheet.getRange(addresses[0]),
      };
    }
    case 'WRITE_TABLE': {
      const colCount = Math.max(
        action.headers.length,
        ...action.rows.map((row) => (Array.isArray(row) ? row.length : 0)),
        1,
      );
      const rowCount = action.headers.length ? 1 + action.rows.length : action.rows.length;
      const address = rangeAddressFromIndexes(0, 0, Math.max(rowCount, 1), colCount);
      return {
        sheetName: action.sheetName,
        address,
        getRange: (sheet) =>
          sheet.getRangeByIndexes(0, 0, Math.max(rowCount, 1), colCount),
      };
    }
    case 'FILL_DOWN':
      return {
        sheetName: action.sheetName,
        address: action.targetRange,
        getRange: (sheet) => sheet.getRange(action.targetRange),
      };
    case 'AUTO_FILL': {
      const start = parseCellAddress(action.startAddress);
      if (!start) return null;
      // Conservative: check the start cell; full dest span depends on live used range.
      return {
        sheetName: action.sheetName,
        address: action.startAddress,
        getRange: (sheet) => sheet.getRange(action.startAddress),
      };
    }
    case 'APPEND_ROW':
      // Target row is derived at execution from used range — guarded inside handler path
      // via a dedicated check after resolving the append index.
      return null;
    case 'COPY_FILTERED_RANGE':
    case 'MOVE_RANGE':
    case 'AGGREGATE_TABLE': {
      const start = parseCellAddress(action.destStartCell);
      if (!start) return null;
      return {
        sheetName: action.destSheet,
        address: action.destStartCell,
        getRange: (sheet) => sheet.getRange(action.destStartCell),
      };
    }
    default:
      return null;
  }
}

function buildOverwriteMessage(
  targetRange: string,
  sample: unknown[][],
): string {
  const flat = sample
    .flat()
    .map((v) => (v == null || v === '' ? null : String(v)))
    .filter((v): v is string => Boolean(v))
    .slice(0, 3);
  const sampleText = flat.length ? ` Existing values include: ${flat.join(', ')}.` : '';
  return (
    `Write blocked: target range ${targetRange} already contains data. ` +
    `This action would overwrite existing values.${sampleText} ` +
    `If you meant to add a new column, use INSERT_COLUMN with position "afterLastColumn" ` +
    `instead of writing into an occupied column. ` +
    `To replace existing content on purpose, the request must explicitly confirm overwrite.`
  );
}

async function assertRangeEmpty(
  sheet: Excel.Worksheet,
  address: string,
  getRange: (sheet: Excel.Worksheet) => Excel.Range,
  ctx: Excel.RequestContext,
): Promise<void> {
  const range = getRange(sheet);
  range.load('values');
  await ctx.sync();
  const values = (range.values ?? []) as unknown[][];
  if (!rangeHasExistingData(values)) return;

  throw new OverwriteGuardError({
    message: buildOverwriteMessage(address, values.slice(0, 3)),
    targetRange: address,
    sampleExistingValues: values.slice(0, 3),
  });
}

/**
 * Pre-write occupancy check. Throws OverwriteGuardError when the target already
 * has data and the action did not set explicitOverwriteConfirmed.
 */
export async function guardAgainstOverwrite(
  action: RichAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  if (!isOverwriteGuardedAction(action)) return;
  if (hasExplicitOverwriteConfirmed(action)) return;

  if (action.type === 'BATCH_SET') {
    const sheet = resolveWorksheet(ctx, action.sheetName);
    for (const op of action.operations) {
      await assertRangeEmpty(
        sheet,
        op.address,
        (s) => s.getRange(op.address),
        ctx,
      );
    }
    return;
  }

  if (action.type === 'APPEND_ROW') {
    const sheet = resolveWorksheet(ctx, action.sheetName);
    const used = sheet.getUsedRange();
    if (used) {
      used.load(['rowCount', 'columnCount']);
      await ctx.sync();
    }
    const targetRow = used ? used.rowCount : 0;
    const colCount = Math.max(action.values.length, used ? used.columnCount : 1, 1);
    const address = rangeAddressFromIndexes(targetRow, 0, 1, colCount);
    await assertRangeEmpty(
      sheet,
      address,
      (s) => s.getRangeByIndexes(targetRow, 0, 1, colCount),
      ctx,
    );
    return;
  }

  if (
    action.type === 'COPY_FILTERED_RANGE' ||
    action.type === 'MOVE_RANGE' ||
    action.type === 'AGGREGATE_TABLE'
  ) {
    // Destination may be a new/empty sheet — only block when the start cell has data.
    const destSheetName = action.destSheet;
    const destStart = action.destStartCell;
    const sheets = ctx.workbook.worksheets;
    let existing: Excel.Worksheet;
    try {
      existing = sheets.getItem(destSheetName);
      existing.load('name');
      await ctx.sync();
    } catch {
      return; // sheet does not exist yet — write is safe
    }
    await assertRangeEmpty(
      existing,
      destStart,
      (s) => s.getRange(destStart),
      ctx,
    );
    return;
  }

  if (action.type === 'AUTO_FILL') {
    // Check cells below/right of start that will be filled (excluding the source cell).
    const sheet = resolveWorksheet(ctx, action.sheetName);
    const start = parseCellAddress(action.startAddress);
    if (!start) return;
    const used = sheet.getUsedRange();
    if (!used) return;
    used.load(['rowCount', 'columnCount']);
    await ctx.sync();

    if (action.direction === 'down') {
      const endRow = (action.endRow ?? used.rowCount) - 1;
      if (endRow <= start.row) return;
      const rowCount = endRow - start.row;
      const address = rangeAddressFromIndexes(start.row + 1, start.col, rowCount, 1);
      await assertRangeEmpty(
        sheet,
        address,
        (s) => s.getRangeByIndexes(start.row + 1, start.col, rowCount, 1),
        ctx,
      );
    } else {
      const endCol = action.endCol ?? used.columnCount - 1;
      if (endCol <= start.col) return;
      const colCount = endCol - start.col;
      const address = rangeAddressFromIndexes(start.row, start.col + 1, 1, colCount);
      await assertRangeEmpty(
        sheet,
        address,
        (s) => s.getRangeByIndexes(start.row, start.col + 1, 1, colCount),
        ctx,
      );
    }
    return;
  }

  const target = resolveWriteTarget(action);
  if (!target) return;

  const sheet = resolveWorksheet(ctx, target.sheetName);
  await assertRangeEmpty(sheet, target.address, target.getRange, ctx);
}

/** Pure helper for tests — describe whether a SheetAction-like write would be guarded. */
export function shouldGuardSheetActionType(type: string): boolean {
  return OVERWRITE_GUARDED_ACTION_TYPES.has(type as RichAction['type']);
}

/** Sample existing header label from values for user-facing messages. */
export function sampleLabelFromValues(values: unknown[][]): string | null {
  for (const row of values) {
    if (!Array.isArray(row)) continue;
    for (const cell of row) {
      if (cell !== '' && cell != null) return String(cell);
    }
  }
  return null;
}
