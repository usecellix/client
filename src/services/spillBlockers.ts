import { columnIndexToLetter, parseCellAddress } from '@/engine/addressUtils';

/* global Excel */

/**
 * `#SPILL!` diagnosis and repair — TASKS.md #321.
 *
 * A spilling formula (Main's live "all bookings" view, TASKS.md #142) shows
 * `#SPILL!` when any cell in the area it needs to fill already has content.
 * The formula itself is correct, so the generic repair — asking the model to
 * rewrite the formula — is the wrong fix, and "1 formula cell(s) returned an
 * error: Main!A19 #SPILL!" tells the user nothing they can act on. What they
 * need is WHICH cells are in the way, and a way to clear them.
 *
 * Excel does not report the intended spill size of a blocked formula, so the
 * area is inferred the way the consolidated view is laid out: as wide as the
 * header row directly above the formula (contiguous labels starting at its
 * column), and down to the sheet's last used row.
 */

/** Widest spill area inspected — keeps the probe bounded on a malformed header. */
const MAX_SPILL_COLUMNS = 50;
/** Deepest spill area inspected. */
const MAX_SPILL_ROWS = 2000;
/** Blocking cells named individually. */
const MAX_BLOCKERS = 20;

export interface SpillBlockage {
  sheet: string;
  /** The formula cell showing #SPILL!, e.g. "A19". */
  anchor: string;
  /** Cells inside its spill area that have content, e.g. ["G19", "I19"]. */
  blockers: string[];
}

function isFilled(cell: unknown): boolean {
  return cell !== '' && cell !== null && cell !== undefined;
}

/** Spill width from the header row above: contiguous labels from the anchor's column. */
export function spillWidthFromHeader(headerCells: unknown[]): number {
  let width = 0;
  for (const cell of headerCells) {
    if (!isFilled(cell) || String(cell).trim() === '') break;
    width += 1;
  }
  return width;
}

/**
 * Cells inside the spill area, other than the anchor itself, that hold a value
 * or a formula. A formula counts even when it currently shows "" — it still
 * occupies the cell (the live G19 `=IF(...,"",...)` case).
 */
export function spillBlockersFromGrid(
  values: unknown[][],
  formulas: unknown[][],
  origin: { row: number; col: number },
): string[] {
  const blockers: string[] = [];
  for (let r = 0; r < values.length; r += 1) {
    for (let c = 0; c < (values[r]?.length ?? 0); c += 1) {
      if (r === 0 && c === 0) continue;
      const formula = formulas[r]?.[c];
      const occupied =
        isFilled(values[r][c]) || (typeof formula === 'string' && formula.startsWith('='));
      if (!occupied) continue;
      blockers.push(`${columnIndexToLetter(origin.col + c)}${origin.row + r + 1}`);
      if (blockers.length >= MAX_BLOCKERS) return blockers;
    }
  }
  return blockers;
}

/**
 * Find what blocks a `#SPILL!` cell. Returns [] when the spill area cannot be
 * inferred (no header row above it) — the caller then falls back to the plain
 * error report rather than guessing.
 */
export async function findSpillBlockers(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  anchorAddress: string,
): Promise<string[]> {
  const anchor = parseCellAddress(anchorAddress);
  if (!anchor || anchor.row === 0) return [];

  const header = sheet.getRangeByIndexes(anchor.row - 1, anchor.col, 1, MAX_SPILL_COLUMNS);
  header.load('values');
  const used = sheet.getUsedRangeOrNullObject();
  used.load(['isNullObject', 'rowIndex', 'rowCount']);
  await ctx.sync();

  const width = spillWidthFromHeader((header.values as unknown[][])?.[0] ?? []);
  if (width === 0) return [];

  const lastUsedRow = used.isNullObject ? anchor.row : used.rowIndex + used.rowCount - 1;
  const height = Math.min(Math.max(lastUsedRow - anchor.row + 1, 1), MAX_SPILL_ROWS);

  const area = sheet.getRangeByIndexes(anchor.row, anchor.col, height, width);
  area.load(['values', 'formulas']);
  await ctx.sync();

  return spillBlockersFromGrid(
    (area.values as unknown[][]) ?? [],
    (area.formulas as unknown[][]) ?? [],
    anchor,
  );
}

/** "G19, I19 and M19" — short enough for a button label. */
export function describeCells(cells: string[], max = 4): string {
  if (cells.length === 0) return '';
  const shown = cells.slice(0, max);
  const rest = cells.length - shown.length;
  if (rest > 0) return `${shown.join(', ')} and ${rest} more`;
  if (shown.length === 1) return shown[0];
  return `${shown.slice(0, -1).join(', ')} and ${shown[shown.length - 1]}`;
}

/**
 * Clear the blocking cells' contents (formatting kept), then read the anchor
 * back. Resolves `true` when the formula now spills cleanly.
 *
 * Only ever runs on an explicit click that names the cells it will clear —
 * this deletes content, so it is the user's call, the same consent rule
 * Accept follows.
 */
export async function clearSpillBlockers(blockage: SpillBlockage): Promise<boolean> {
  let resolved = false;
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(blockage.sheet);
    for (const cell of blockage.blockers) {
      sheet.getRange(cell).clear(Excel.ClearApplyTo.contents);
    }
    await ctx.sync();

    const anchor = sheet.getRange(blockage.anchor);
    anchor.load('values');
    await ctx.sync();
    const value = (anchor.values as unknown[][])?.[0]?.[0];
    resolved = !(typeof value === 'string' && /^#SPILL!?$/i.test(value.trim()));
  });
  return resolved;
}

/** The blocked spills a read-back found, ready to offer as a one-click clear. */
export function buildSpillBlockages(
  mismatches: Array<{ sheet: string; cell: string; spillBlockers?: string[] }> | undefined,
): SpillBlockage[] {
  return (mismatches ?? [])
    .filter((m) => m.spillBlockers && m.spillBlockers.length > 0)
    .map((m) => ({ sheet: m.sheet, anchor: m.cell, blockers: [...(m.spillBlockers ?? [])] }));
}
