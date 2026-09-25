import { RichAction } from '@/action.types';
import { columnIndexToLetter, parseCellAddress, stripSheetPrefix } from './addressUtils';
import { resolveWorksheet } from './handlers/resolveWorksheet';
import { findMatchingRowOffsets, resolveFilterColumnIndex } from './rangeFilter';

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
  'SET_MATCHING_ROWS',
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
  actionType?: RichAction['type'],
): string {
  const flat = sample
    .flat()
    .map((v) => (v == null || v === '' ? null : String(v)))
    .filter((v): v is string => Boolean(v))
    .slice(0, 3);
  const sampleText = flat.length ? ` Existing values include: ${flat.join(', ')}.` : '';
  const isDestWrite =
    actionType === 'COPY_FILTERED_RANGE' ||
    actionType === 'MOVE_RANGE' ||
    actionType === 'AGGREGATE_TABLE';

  if (isDestWrite) {
    return (
      `Write blocked: destination ${targetRange} already contains data.` +
      `${sampleText} ` +
      `Clear the destination sheet first, choose an empty start cell, or confirm overwrite ` +
      `if you meant to replace the existing contents.`
    );
  }

  return (
    `Write blocked: target range ${targetRange} already contains data. ` +
    `This action would overwrite existing values.${sampleText} ` +
    `If you meant to add a new column, use INSERT_COLUMN with position "afterLastColumn" ` +
    `instead of writing into an occupied column. ` +
    `To replace existing content on purpose, the request must explicitly confirm overwrite.`
  );
}

/** What a single-cell write intends the cell to hold. */
export interface ExpectedCellContent {
  formula?: string;
  value?: unknown;
}

function normalizeFormula(formula: string): string {
  return formula.replace(/\s+/g, '').toUpperCase();
}

/**
 * A cell that already holds exactly what this write would put there is not an
 * overwrite — rewriting it changes nothing. TASKS.md #319: a step blocked
 * partway (its earlier writes already landed) was then refused on every retry
 * by its OWN earlier writes ("A1 already contains data: Payments Dashboard").
 *
 * Compared against `formulas`, not `values`: a formula returning "" reads as
 * empty in `values`, and a formula's value says nothing about whether it is the
 * same formula.
 */
export function cellAlreadyHolds(
  current: { value: unknown; formula: unknown },
  expected: ExpectedCellContent,
): boolean {
  const currentFormula = typeof current.formula === 'string' ? current.formula : '';
  if (typeof expected.formula === 'string' && expected.formula.trim()) {
    return (
      currentFormula.startsWith('=') &&
      normalizeFormula(currentFormula) === normalizeFormula(expected.formula)
    );
  }
  if (expected.value === undefined || expected.value === null) return false;
  if (currentFormula.startsWith('=')) return false;
  return String(current.value).trim() === String(expected.value).trim();
}

async function assertRangeEmpty(
  sheet: Excel.Worksheet,
  address: string,
  getRange: (sheet: Excel.Worksheet) => Excel.Range,
  ctx: Excel.RequestContext,
  actionType?: RichAction['type'],
  expected?: ExpectedCellContent,
): Promise<void> {
  const range = getRange(sheet);
  range.load(expected ? ['values', 'formulas'] : 'values');
  await ctx.sync();
  const values = (range.values ?? []) as unknown[][];
  if (!rangeHasExistingData(values)) {
    // A formula returning "" reads as empty in `values` but is still content.
    if (!expected) return;
    const formula = (range.formulas as unknown[][] | undefined)?.[0]?.[0];
    if (!(typeof formula === 'string' && formula.startsWith('='))) return;
    if (cellAlreadyHolds({ value: values[0]?.[0], formula }, expected)) return;
  } else if (
    expected &&
    values.length === 1 &&
    values[0]?.length === 1 &&
    cellAlreadyHolds(
      { value: values[0][0], formula: (range.formulas as unknown[][] | undefined)?.[0]?.[0] },
      expected,
    )
  ) {
    return;
  }

  throw new OverwriteGuardError({
    message: buildOverwriteMessage(address, values.slice(0, 3), actionType),
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
    // The backend now rejects a BATCH_SET with no operations array before it
    // ever reaches here (normalize-executor-output.util.ts), but this guarded
    // check stays as defense in depth — an unguarded `for...of` on a missing
    // array here previously threw "action.operations is not iterable" and
    // aborted the entire apply, including every other already-verified action
    // in the same batch.
    if (!Array.isArray(action.operations) || action.operations.length === 0) return;
    const sheet = resolveWorksheet(ctx, action.sheetName);
    for (const op of action.operations) {
      await assertRangeEmpty(
        sheet,
        op.address,
        (s) => s.getRange(op.address),
        ctx,
        undefined,
        { formula: op.formula, value: op.value },
      );
    }
    return;
  }

  if (action.type === 'APPEND_ROW') {
    const sheet = resolveWorksheet(ctx, action.sheetName);
    // `getUsedRange()` throws ItemNotFound on a sheet with literally nothing on
    // it — the exact state of a sheet `ADD_SHEET` just created. That throw
    // takes out the whole batch (worksheet.handler.ts's documented failure
    // mode), which is precisely what an ADD_ROW-shaped header write into a
    // brand-new sheet hits. `getUsedRangeOrNullObject` is the null-safe
    // sibling and needs the standard load-then-sync-then-branch shape.
    // TASKS.md #146.
    const used = sheet.getUsedRangeOrNullObject();
    used.load(['rowCount', 'columnCount', 'isNullObject']);
    await ctx.sync();
    const exists = !used.isNullObject;
    const targetRow = exists ? used.rowCount : 0;
    const colCount = Math.max(action.values.length, exists ? used.columnCount : 1, 1);
    const address = rangeAddressFromIndexes(targetRow, 0, 1, colCount);
    await assertRangeEmpty(
      sheet,
      address,
      (s) => s.getRangeByIndexes(targetRow, 0, 1, colCount),
      ctx,
    );
    return;
  }

  if (action.type === 'SET_MATCHING_ROWS') {
    // Blanking cells is a clear operation — allow replacing existing values.
    if (
      action.value === null ||
      action.value === undefined ||
      String(action.value).trim() === ''
    ) {
      return;
    }

    const sheet = resolveWorksheet(ctx, action.sheetName);
    const localRange = stripSheetPrefix(action.range);
    const range = sheet.getRange(localRange);
    range.load(['values', 'rowIndex', 'columnIndex']);
    await ctx.sync();
    const rows = (range.values ?? []) as unknown[][];
    if (rows.length === 0) return;
    const hasHeaders = action.hasHeaders !== false;
    const headerRow = hasHeaders ? rows[0] : null;
    if (!headerRow) return;
    let targetIdx: number;
    try {
      targetIdx = resolveFilterColumnIndex(headerRow, action.targetColumn);
    } catch {
      return;
    }
    const offsets = action.filter
      ? findMatchingRowOffsets(rows, hasHeaders, action.filter)
      : Array.from({ length: Math.max(rows.length - (hasHeaders ? 1 : 0), 0) }, (_, i) =>
          hasHeaders ? i + 1 : i,
        );
    const occupiedSamples: unknown[][] = [];
    for (const offset of offsets) {
      const cellVal = rows[offset]?.[targetIdx];
      if (rangeHasExistingData([[cellVal]])) {
        occupiedSamples.push([cellVal]);
      }
    }
    if (occupiedSamples.length === 0) return;
    const sampleAddress = `${columnIndexToLetter(range.columnIndex + targetIdx)}${
      range.rowIndex + (offsets[0] ?? 0) + 1
    }`;
    throw new OverwriteGuardError({
      message: buildOverwriteMessage(sampleAddress, occupiedSamples.slice(0, 3), action.type),
      targetRange: sampleAddress,
      sampleExistingValues: occupiedSamples.slice(0, 3),
    });
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
      action.type,
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

  const expected: ExpectedCellContent | undefined =
    action.type === 'SET_FORMULA'
      ? { formula: action.formula }
      : action.type === 'SET_CELL'
        ? { value: action.value }
        : undefined;
  const sheet = resolveWorksheet(ctx, target.sheetName);
  await assertRangeEmpty(sheet, target.address, target.getRange, ctx, action.type, expected);
}

/**
 * Action types that neither move nor remove existing cell content, so a later
 * write in the same card still targets the cells the workbook holds NOW. The
 * pre-flight stops at the first action NOT in this list: after an insert,
 * delete, clear, sort or move, a later write may legitimately target cells that
 * action emptied, and judging it against today's workbook would wrongly block a
 * valid card.
 */
const PREFLIGHT_SAFE_TYPES = new Set<string>([
  'ADD_SHEET',
  'SET_CELL',
  'SET_FORMULA',
  'BATCH_SET',
  'SET_RANGE_VALUES',
  'WRITE_TABLE',
  'FILL_DOWN',
  'AUTO_FILL',
  'APPEND_ROW',
  'FORMAT_RANGE',
  'CONDITIONAL_FORMAT',
  'DATA_VALIDATION',
  'SET_COLUMN_WIDTH',
  'SET_ROW_HEIGHT',
  'AUTOFIT_COLUMNS',
  'HIDE_GRIDLINES',
  'FREEZE_PANES',
  'SET_SHEET_COLOR',
  'SET_ZOOM',
  'HIGHLIGHT_CELL',
  'CREATE_CHART',
  'CREATE_TABLE',
  'DEFINE_NAMED_RANGE',
  'ADD_COMMENT',
]);

/** The sheet a guarded action's occupancy check reads. */
function guardedSheetOf(action: RichAction): string | undefined {
  const record = action as unknown as { destSheet?: unknown; sheetName?: unknown };
  const name =
    action.type === 'COPY_FILTERED_RANGE' ||
    action.type === 'MOVE_RANGE' ||
    action.type === 'AGGREGATE_TABLE'
      ? record.destSheet
      : record.sheetName;
  return typeof name === 'string' && name.trim() ? name : undefined;
}

/**
 * Run the overwrite guard over a whole card BEFORE anything is written —
 * TASKS.md #319.
 *
 * The per-action guard runs inside the write loop, and its own `ctx.sync()`
 * commits every write queued before it. So a block at action 6 left actions 1–5
 * applied, the card still pending, and every retry refused by those earlier
 * writes. A live step 6 was blocked at A19 after its title, KPIs and chart had
 * landed, then blocked at A1 on each retry.
 *
 * Checking first means a block leaves the workbook untouched. Only
 * `OverwriteGuardError` propagates. Any other failure here (a sheet that
 * resolves oddly, a host quirk) is left for the write loop, which reports it
 * exactly as before.
 */
export async function preflightOverwriteGuard(
  actions: RichAction[],
  ctx: Excel.RequestContext,
): Promise<void> {
  const createdHere = new Set<string>();

  for (const action of actions) {
    if (!PREFLIGHT_SAFE_TYPES.has(action.type)) return;

    if (action.type === 'ADD_SHEET') {
      const record = action as unknown as { name?: unknown; sheetName?: unknown };
      const name = String(record.name ?? record.sheetName ?? '').trim().toLowerCase();
      if (name) createdHere.add(name);
      continue;
    }

    if (!isOverwriteGuardedAction(action) || hasExplicitOverwriteConfirmed(action)) continue;

    const sheetName = guardedSheetOf(action);
    if (!sheetName || createdHere.has(sheetName.toLowerCase())) continue;

    try {
      const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      sheet.load('isNullObject');
      await ctx.sync();
      if (sheet.isNullObject) continue;
      await guardAgainstOverwrite(action, ctx);
    } catch (error) {
      if (isOverwriteGuardError(error)) throw error;
    }
  }
}

/**
 * When a proposal creates a sheet and then COPY/MOVE/AGGREGATE into it, the write
 * is the intended fill of that destination — not accidental overwrite of user data.
 * Mark those dest writes so Accept can complete even if a prior attempt already
 * populated the sheet (or ADD_SHEET no-op'd because the tab already existed).
 */
export function annotateDestOverwriteForCreatedSheets<T extends { type: string; name?: string; sheetName?: string; destSheet?: string; explicitOverwriteConfirmed?: boolean }>(
  actions: T[],
): T[] {
  const created = new Set<string>();
  for (const action of actions) {
    if (action.type === 'ADD_SHEET' || action.type === 'CREATE_SHEET') {
      const name = String(action.name ?? action.sheetName ?? '').trim().toLowerCase();
      if (name && !/^sheet\d*$/i.test(name)) created.add(name);
    }
  }
  if (created.size === 0) return actions;

  return actions.map((action) => {
    if (
      action.type !== 'COPY_FILTERED_RANGE' &&
      action.type !== 'MOVE_RANGE' &&
      action.type !== 'AGGREGATE_TABLE'
    ) {
      return action;
    }
    const dest = String(action.destSheet ?? '').trim().toLowerCase();
    if (!dest || !created.has(dest)) return action;
    return { ...action, explicitOverwriteConfirmed: true };
  });
}

/** Drop placeholder Sheet2/SheetN creates when a real named sheet is also in the batch. */
export function pruneSpuriousAddSheets<T extends { type: string; name?: string; sheetName?: string; destSheet?: string }>(
  actions: T[],
): T[] {
  if (actions.length < 2) return actions;
  const namedCreates = new Set(
    actions
      .filter((a) => a.type === 'ADD_SHEET' || a.type === 'CREATE_SHEET')
      .map((a) => String(a.name ?? a.sheetName ?? '').trim().toLowerCase())
      .filter((n) => n && !/^sheet\d*$/i.test(n)),
  );
  if (namedCreates.size === 0) return actions;

  const destSheets = new Set(
    actions
      .filter(
        (a) =>
          a.type === 'COPY_FILTERED_RANGE' ||
          a.type === 'MOVE_RANGE' ||
          a.type === 'AGGREGATE_TABLE',
      )
      .map((a) => String(a.destSheet ?? '').trim().toLowerCase())
      .filter(Boolean),
  );

  return actions.filter((action) => {
    if (action.type !== 'ADD_SHEET' && action.type !== 'CREATE_SHEET') return true;
    const name = String(action.name ?? action.sheetName ?? '').trim();
    if (!name) return false;
    if (/^sheet\d*$/i.test(name) && (namedCreates.size > 0 || destSheets.size > 0)) {
      return false;
    }
    return true;
  });
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
