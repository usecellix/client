import { SheetAction } from '@/types/sheet-actions';

/* global Excel */

/**
 * Per-sheet facts the header guard needs, keyed by sheet name.
 *
 * `sanitizeActions` exists to stop an LLM writing data over a user's real
 * column headers. That question — "does this sheet already have headers?" —
 * is a fact about a *specific live sheet*, and it must be answered per sheet.
 *
 * It used to be inferred from the shape of the action list instead
 * (`detectPopulateEmptySheet`: "does this batch contain both a row-0 SET_CELL
 * and an ADD_ROW?"). That inference is wrong in both directions, and its
 * false-negative destroyed every multi-sheet build: a batch that creates 12
 * month sheets and writes their A1:J1 headers contains no ADD_ROW, so the
 * guard concluded the sheets were populated and rewrote 121 header cells into
 * a single row on the active tab. See TASKS.md #137.
 */
export interface SheetGuardState {
  /** True when row 1 of the live sheet already holds a value. */
  hasHeaderRow: boolean;
  /** False when the sheet does not exist yet (an earlier wave will create it). */
  exists: boolean;
}

export type SheetGuardStates = Map<string, SheetGuardState>;

/** Columns of row 1 probed per sheet — enough to detect a header row. */
const HEADER_PROBE_COLUMNS = 64;

const ACTIVE_SHEET_KEY = '';

/**
 * Normalized lookup key for an action's target sheet. Empty string means
 * "whatever sheet is active", which is how `resolveWorksheet` treats a
 * missing `sheetName`.
 */
export function sheetKeyOf(action: SheetAction): string {
  const name = String(action.sheetName ?? (action as { destSheet?: string }).destSheet ?? '').trim();
  return name.toLowerCase();
}

/** Distinct target sheet names in a batch, preserving original casing. */
export function targetSheetNames(actions: SheetAction[]): string[] {
  const byKey = new Map<string, string>();
  for (const action of actions) {
    const name = String(action.sheetName ?? (action as { destSheet?: string }).destSheet ?? '').trim();
    if (!name) continue;
    const key = name.toLowerCase();
    if (!byKey.has(key)) byKey.set(key, name);
  }
  return [...byKey.values()];
}

/**
 * Sheets this batch creates itself. They are empty by construction, so a
 * row-1 write into one is a header write, never an overwrite — this is the
 * same created-sheet reasoning `annotateDestOverwriteForCreatedSheets` already
 * applies to COPY/MOVE destinations.
 *
 * Note this only covers creates in the *same* batch. Staged accept waves put
 * the ADD_SHEETs in wave 1 and the writes in wave 2, so wave 2 sees no creates
 * at all — which is exactly why the live probe below is the primary source of
 * truth and this is only a supplement.
 */
export function sheetsCreatedInBatch(actions: SheetAction[]): Set<string> {
  const created = new Set<string>();
  for (const action of actions) {
    if (action.type !== 'ADD_SHEET' && action.type !== 'CREATE_SHEET') continue;
    const name = String(action.name ?? action.sheetName ?? '').trim();
    if (name) created.add(name.toLowerCase());
  }
  return created;
}

/**
 * Read row 1 of every sheet this batch targets, in one Office.js round trip.
 *
 * Uses `getItemOrNullObject` so a sheet that does not exist yet cannot poison
 * the request context — the same failure mode documented in
 * `worksheet.handler.ts`, where a lookup for a missing sheet stayed queued and
 * failed the next `ctx.sync()` for the whole batch.
 */
export async function probeSheetGuardStates(
  actions: SheetAction[],
): Promise<SheetGuardStates> {
  const states: SheetGuardStates = new Map();
  const names = targetSheetNames(actions);
  const createdHere = sheetsCreatedInBatch(actions);

  // A sheet created by this very batch is empty by construction — no probe needed
  // (and probing it would read the pre-creation state anyway).
  for (const key of createdHere) {
    states.set(key, { hasHeaderRow: false, exists: false });
  }

  const toProbe = names.filter((name) => !createdHere.has(name.toLowerCase()));
  const needsActiveSheet = actions.some((action) => sheetKeyOf(action) === ACTIVE_SHEET_KEY);

  if (toProbe.length === 0 && !needsActiveSheet) return states;

  await Excel.run(async (ctx) => {
    // Two-phase, not one. Chaining `.getRangeByIndexes()` off a
    // `getItemOrNullObject()` result and loading both in the SAME sync is
    // exactly the failure mode `worksheet.handler.ts` documents: Office.js
    // resolves the range request against the (possibly still-nonexistent)
    // parent before this call learns `isNullObject`, and a missing sheet
    // throws "The requested resource doesn't exist." — for the WHOLE sync,
    // not just that one probe. `previewManager.ts`'s `render()` already gets
    // this right (load `isNullObject`, sync, THEN branch); this probe did not,
    // and it broke a live run: every prompt failed before a single sheet was
    // created. See TASKS.md #145.
    const sheets = new Map<string, Excel.Worksheet>();

    for (const name of toProbe) {
      const sheet = ctx.workbook.worksheets.getItemOrNullObject(name);
      sheet.load('isNullObject');
      sheets.set(name.toLowerCase(), sheet);
    }

    let activeSheet: Excel.Worksheet | undefined;
    if (needsActiveSheet) {
      activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
    }

    // Phase 1: learn which sheets exist. Nothing chained off a possibly-null
    // object yet.
    await ctx.sync();

    const ranges: { key: string; range: Excel.Range }[] = [];
    for (const [key, sheet] of sheets) {
      if (sheet.isNullObject) {
        states.set(key, { hasHeaderRow: false, exists: false });
        continue;
      }
      const range = sheet.getRangeByIndexes(0, 0, 1, HEADER_PROBE_COLUMNS);
      range.load('values');
      ranges.push({ key, range });
    }
    if (activeSheet) {
      const range = activeSheet.getRangeByIndexes(0, 0, 1, HEADER_PROBE_COLUMNS);
      range.load('values');
      ranges.push({ key: ACTIVE_SHEET_KEY, range });
    }

    if (ranges.length === 0) return;

    // Phase 2: only sheets confirmed to exist reach here.
    await ctx.sync();

    for (const { key, range } of ranges) {
      const values = (range.values ?? []) as unknown[][];
      states.set(key, { hasHeaderRow: rowHasContent(values[0]), exists: true });
    }
  });

  return states;
}

function rowHasContent(row: unknown[] | undefined): boolean {
  if (!Array.isArray(row)) return false;
  return row.some((cell) => cell !== '' && cell != null);
}

/**
 * Best-effort probe. A probe failure must never block an apply — fall back to
 * an empty map, which makes `sanitizeActions` use its legacy single-sheet
 * `layout` behaviour rather than silently guessing per sheet.
 */
export async function probeSheetGuardStatesSafe(
  actions: SheetAction[],
): Promise<SheetGuardStates> {
  try {
    return await probeSheetGuardStates(actions);
  } catch (error) {
    console.warn('[Cellix] Header-row probe failed; falling back to batch-shape guard:', error);
    return new Map();
  }
}
