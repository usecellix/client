import { OutcomeVerification, OutcomeMismatch } from './outcomeVerifier';

/**
 * Turn a failed read-back into a repair the user can actually run — TASKS.md #168.
 *
 * ## Why this exists
 *
 * TASKS.md #150 gave the system the ability to notice that a formula it just
 * wrote came back as `#REF!`. It could report that and nothing more. Shortcut's
 * own transcript (planning log part 6) shows the move that follows: it saw the
 * `#REF!` in its consolidation formula, wrote small test formulas to a scratch
 * cell to isolate the cause, found a `LET` variable name colliding with a real
 * range reference, and fixed the live formula. Ours stopped at "Applied, but 1
 * formula cell returned an error" and left the broken formula sitting in the
 * user's dashboard.
 *
 * Full automatic mid-run repair needs the resumable loop (TASKS.md #153/#154):
 * the run is over by the time the read-back happens, so nothing can re-plan.
 * But the *information* needed to repair is complete right here — the cell, the
 * error, and the formula that produced it — and a follow-up turn is a path the
 * architecture already has. So this builds the repair as a normal next message,
 * which needs no new endpoint, no run persistence and no SSE contract change.
 * It is a smaller thing than #154 and it is not a substitute for it; it closes
 * the "reports but cannot act" gap using only what exists today.
 *
 * ## What it deliberately will NOT repair
 *
 * Only Excel error literals. A plain value mismatch is excluded on purpose:
 * Excel's own smart-entry re-parses values on write (`12-09-26` becoming a
 * date, a long number losing precision), so a mismatch there frequently means
 * "Excel did something reasonable that we failed to predict", and asking the
 * model to force the predicted value back would fight the host rather than fix
 * a bug. Those still get reported by `describeOutcome`; they just do not
 * generate a repair. A missing sheet is likewise excluded — that is a
 * structural failure needing a rebuild, not a formula fix.
 */

export interface RepairRequest {
  /** A natural-language follow-up message, ready to send as the next turn. */
  prompt: string;
  /** How many erroring cells triggered it. */
  cellCount: number;
  /** Distinct error literals seen, e.g. ['#REF!', '#NAME?']. */
  errors: string[];
}

/** Cap on cells named individually — a repair prompt must stay readable. */
const MAX_CELLS_NAMED = 8;

function normalizeError(actual: string): string {
  const trimmed = String(actual ?? '').trim().toUpperCase();
  return trimmed.endsWith('!') || trimmed.endsWith('?') ? trimmed : `${trimmed}!`;
}

/**
 * Group by the error literal AND the formula, because a build that writes the
 * same broken formula down twelve month rows has ONE bug, not twelve. Naming
 * all twelve would spend the prompt's budget restating a single mistake.
 */
function groupFailures(mismatches: OutcomeMismatch[]): Map<string, OutcomeMismatch[]> {
  const groups = new Map<string, OutcomeMismatch[]>();
  for (const mismatch of mismatches) {
    const key = `${normalizeError(mismatch.actual)}|${mismatch.formula ?? mismatch.expected ?? ''}`;
    const existing = groups.get(key);
    if (existing) existing.push(mismatch);
    else groups.set(key, [mismatch]);
  }
  return groups;
}

export function buildRepairRequest(result: OutcomeVerification): RepairRequest | null {
  if (!result || result.skipped) return null;

  const failures = (result.mismatches ?? []).filter(
    (mismatch) => mismatch.isFormulaError && mismatch.actual !== '(sheet missing)',
  );
  if (failures.length === 0) return null;

  const groups = groupFailures(failures);
  const errors = [...new Set(failures.map((f) => normalizeError(f.actual)))];

  const lines: string[] = [];
  for (const [, members] of groups) {
    const first = members[0];
    const where = members
      .slice(0, MAX_CELLS_NAMED)
      .map((m) => `${m.sheet}!${m.cell}`)
      .join(', ');
    const more = members.length > MAX_CELLS_NAMED ? ` and ${members.length - MAX_CELLS_NAMED} more cell(s)` : '';
    const formula = first.formula ? ` The formula written there was: ${first.formula}` : '';
    lines.push(`- ${where}${more} returned ${normalizeError(first.actual)}.${formula}`);
  }

  const prompt = [
    'The changes were applied, but these cells came back as Excel errors. Please fix them.',
    '',
    ...lines,
    '',
    'Work out why each formula fails before rewriting it — check that every sheet and range it references exists and is spelled correctly, that no LET/LAMBDA variable name collides with a real range reference, and that any dynamic-array function used is supported here. Rewrite only the failing cells; leave everything else exactly as it is.',
  ].join('\n');

  return { prompt, cellCount: failures.length, errors };
}
