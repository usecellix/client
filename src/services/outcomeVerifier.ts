import { CellChange } from '@/types/changeSet';

/* global Excel */

/**
 * Post-apply outcome verification — TASKS.md #150.
 *
 * `CODEBASE_ANALYSIS.md` §3.15 calls this "the largest structural gap" in the
 * product, and `ARCHITECTURE.md` AD-3 names it as its own required complement:
 * *"a post-apply read-back on the frontend (the only component with live
 * access, per AD-1)"*. Until now Cellix verified **intent** at three points and
 * **outcome** at none:
 *
 *   FormulaValidator  -> formula text        vs static rules
 *   Verifier checkers -> emitted actions     vs estimatedActions (same plan!)
 *   overwriteGuard    -> target occupancy    vs pre-write state
 *   (nothing)         -> the workbook after  vs nothing
 *
 * Every user-visible failure in `COMPETITIVE_STUDY_SHORTCUT.md` lived in that
 * last row: `Main` becoming `Main 2`, Accept applying nothing, a sort flattening
 * dates. All three produce well-formed actions and a wrong workbook, so no
 * amount of extra pre-write checking could ever have caught them.
 *
 * Deliberately **deterministic — no LLM call**. Comparing observed cells to the
 * ChangeSet's own recorded `after` values is a diff; a model would be slower,
 * costlier and less reliable at it. This is a fifth checker, not a fourth
 * agent, and it adds zero tokens to the happy path.
 *
 * Honest limit, stated up front: this answers *"did what I asked for actually
 * happen?"*. It cannot answer *"did I ask for the right things?"* — an
 * under-planned request (TASKS.md #82's truncation class) has nothing missing
 * from the ChangeSet to compare against, so it verifies clean. Closing that
 * needs a different check against the original request, not against the plan.
 */

export interface OutcomeMismatch {
  sheet: string;
  cell: string;
  expected: string;
  actual: string;
  /** True when the cell holds an Excel error (#REF!, #NAME?, ...). */
  isFormulaError: boolean;
  /**
   * The formula the ChangeSet wrote here, when it wrote one.
   *
   * `expected` holds the recorded `after` VALUE, which for a formula cell is
   * not the formula text — so without this a repair has the symptom and not
   * the cause. TASKS.md #168.
   */
  formula?: string;
}

export interface OutcomeVerification {
  /** Cells that were read back and matched. */
  verified: number;
  /** Cells whose live value disagrees with the ChangeSet's `after`. */
  mismatches: OutcomeMismatch[];
  /** Cells that could not be read at all (missing sheet, unreadable range). */
  unreadable: number;
  /** True when the read-back itself could not run — verification is unknown, not clean. */
  skipped: boolean;
  skipReason?: string;
}

/** Excel error literals that mean a written formula did not evaluate. */
const FORMULA_ERROR_PATTERN = /^#(REF|NAME|VALUE|DIV\/0|N\/A|NULL|NUM|SPILL|CALC|GETTING_DATA)[!?]?$/i;

export function isFormulaError(value: unknown): boolean {
  return typeof value === 'string' && FORMULA_ERROR_PATTERN.test(value.trim());
}

/**
 * Compare a written value against what the ChangeSet said it would be.
 *
 * Intentionally lenient about representation, strict about meaning:
 *   - numeric equality wins over string equality (`0` vs `"0"`), because the
 *     ChangeSet records pre-format values while Excel returns typed ones;
 *   - blank/null/'' are all "empty";
 *   - a cell holding a formula returns its COMPUTED value, which will almost
 *     never equal the formula text — those are excluded from comparison by the
 *     caller and only checked for error literals.
 */
export function valuesAgree(expected: unknown, actual: unknown): boolean {
  const e = expected ?? '';
  const a = actual ?? '';
  if (e === a) return true;

  const eStr = String(e).trim();
  const aStr = String(a).trim();
  if (eStr === aStr) return true;
  if (eStr === '' && aStr === '') return true;

  const eNum = Number(eStr);
  const aNum = Number(aStr);
  if (eStr !== '' && aStr !== '' && !Number.isNaN(eNum) && !Number.isNaN(aNum)) {
    // Tolerate float representation drift; exact for integers.
    return Math.abs(eNum - aNum) < 1e-9;
  }

  return false;
}

/** Group changes by sheet, preserving order. */
function groupBySheet(changes: CellChange[]): Map<string, CellChange[]> {
  const bySheet = new Map<string, CellChange[]>();
  for (const change of changes) {
    const sheet = String(change.sheet ?? '').trim();
    if (!sheet || !change.cell) continue;
    const list = bySheet.get(sheet);
    if (list) list.push(change);
    else bySheet.set(sheet, [change]);
  }
  return bySheet;
}

export interface OutcomeVerificationOptions {
  /**
   * Sheet names the plan intended to create, e.g. every ADD_SHEET/CREATE_SHEET
   * action's requested name. Checked directly against `workbook.worksheets`
   * independent of `changes`, because the `Main` -> `Main 2` failure mode
   * (COMPETITIVE_STUDY_SHORTCUT.md:71) means the ChangeSet's own cell rows say
   * "Main" while the write actually landed on a different, unlisted sheet —
   * reading back "Main" alone verifies clean even though the plan's sheet was
   * never created. This is the structural check that closes that gap.
   */
  expectedSheetNames?: string[];
}

/**
 * Read back every cell this ChangeSet claims to have written and compare.
 *
 * Reads are issued per cell but batched into one `ctx.sync()` per sheet, so a
 * 190-cell build costs ~13 round trips rather than 190.
 */
export async function verifyAppliedOutcome(
  changes: CellChange[],
  options?: OutcomeVerificationOptions,
): Promise<OutcomeVerification> {
  const empty: OutcomeVerification = {
    verified: 0,
    mismatches: [],
    unreadable: 0,
    skipped: false,
  };

  const expectedSheetNames = [...new Set(options?.expectedSheetNames ?? [])].filter(Boolean);

  if (!changes?.length && expectedSheetNames.length === 0) return empty;

  if (typeof Excel === 'undefined' || typeof Excel.run !== 'function') {
    return { ...empty, skipped: true, skipReason: 'Office.js unavailable' };
  }

  const bySheet = groupBySheet(changes ?? []);
  if (bySheet.size === 0 && expectedSheetNames.length === 0) return empty;

  const mismatches: OutcomeMismatch[] = [];
  let verified = 0;
  let unreadable = 0;

  await Excel.run(async (ctx) => {
    // Structural manifest check, ahead of the per-cell loop: does every sheet
    // the plan meant to create actually exist under that name? A batch that
    // creates "Main" but lands on "Main 2" has no cell in `changes` addressed
    // to "Main 2" at all, so the per-cell loop below has nothing to disagree
    // with — this is the only check that looks at the sheet *set*, not cells.
    if (expectedSheetNames.length > 0) {
      const sheets = expectedSheetNames.map((name) => ({
        name,
        item: ctx.workbook.worksheets.getItemOrNullObject(name),
      }));
      sheets.forEach(({ item }) => item.load('isNullObject'));
      await ctx.sync();

      for (const { name, item } of sheets) {
        if (item.isNullObject) {
          mismatches.push({
            sheet: name,
            cell: '(sheet)',
            expected: name,
            actual: '(sheet not created)',
            isFormulaError: false,
          });
        }
      }
    }

    for (const [sheetName, sheetChanges] of bySheet) {
      // Two-phase, always: confirm the sheet exists before chaining a range off
      // it. Doing both in one sync is the exact failure mode that produced
      // TASKS.md #145/#146/#147.
      const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      sheet.load('isNullObject');
      await ctx.sync();

      if (sheet.isNullObject) {
        // The ChangeSet says we wrote here and the sheet is not there at all.
        // That is the strongest possible outcome failure, not an unreadable cell.
        for (const change of sheetChanges) {
          mismatches.push({
            sheet: sheetName,
            cell: change.cell,
            expected: String(change.after ?? ''),
            actual: '(sheet missing)',
            isFormulaError: false,
          });
        }
        continue;
      }

      const probes = sheetChanges.map((change) => {
        const range = sheet.getRange(change.cell);
        range.load(['values', 'valueTypes']);
        return { change, range };
      });

      await ctx.sync();

      for (const { change, range } of probes) {
        const actual = (range.values as unknown[][] | undefined)?.[0]?.[0];

        if (actual === undefined) {
          unreadable += 1;
          continue;
        }

        // An Excel error literal is always a failure, formula cell or not —
        // this is how #REF!/#NAME? from a bad consolidation formula surfaces
        // instead of silently sitting in the sheet (TASKS.md #152/#154).
        if (isFormulaError(actual)) {
          mismatches.push({
            sheet: sheetName,
            cell: change.cell,
            expected: String(change.after ?? ''),
            actual: String(actual),
            isFormulaError: true,
            ...(change.formula ? { formula: String(change.formula) } : {}),
          });
          continue;
        }

        // A formula cell returns its computed result, which will not equal the
        // formula text the ChangeSet recorded. Having confirmed it is not an
        // error, that is a pass.
        if (change.formula) {
          verified += 1;
          continue;
        }

        if (valuesAgree(change.after, actual)) {
          verified += 1;
        } else {
          mismatches.push({
            sheet: sheetName,
            cell: change.cell,
            expected: String(change.after ?? ''),
            actual: String(actual),
            isFormulaError: false,
          });
        }
      }
    }
  });

  return { verified, mismatches, unreadable, skipped: false };
}

/** Never let a verification failure break an apply that already succeeded. */
export async function verifyAppliedOutcomeSafe(
  changes: CellChange[],
  options?: OutcomeVerificationOptions,
): Promise<OutcomeVerification> {
  try {
    return await verifyAppliedOutcome(changes, options);
  } catch (error) {
    const reason = error instanceof Error ? error.message : 'read-back failed';
    console.warn('[Cellix] Post-apply verification could not run:', error);
    return { verified: 0, mismatches: [], unreadable: 0, skipped: true, skipReason: reason };
  }
}

/**
 * One honest sentence for the UI.
 *
 * `null` means "nothing worth saying" — a clean verification is silent. The
 * §3.7 rule this follows: never report success the system has not established,
 * and never stay silent about incompleteness it HAS established.
 */
export function describeOutcome(result: OutcomeVerification): string | null {
  if (result.skipped) return null;
  if (result.mismatches.length === 0) return null;

  const errors = result.mismatches.filter((m) => m.isFormulaError);
  const missingSheets = new Set(
    result.mismatches.filter((m) => m.actual === '(sheet missing)').map((m) => m.sheet),
  );
  const notCreatedSheets = new Set(
    result.mismatches.filter((m) => m.actual === '(sheet not created)').map((m) => m.sheet),
  );
  const renamedSheets = result.mismatches.filter(
    (m) => m.cell === '(sheet)' && m.actual !== '(sheet missing)' && m.actual !== '(sheet not created)',
  );

  if (missingSheets.size > 0) {
    const names = [...missingSheets].slice(0, 3).join(', ');
    const more = missingSheets.size > 3 ? ` +${missingSheets.size - 3} more` : '';
    return `Applied, but ${missingSheets.size} sheet(s) could not be found afterwards: ${names}${more}. The workbook may not match what was proposed.`;
  }

  if (notCreatedSheets.size > 0) {
    const names = [...notCreatedSheets].slice(0, 3).join(', ');
    const more = notCreatedSheets.size > 3 ? ` +${notCreatedSheets.size - 3} more` : '';
    return `Applied, but ${notCreatedSheets.size} sheet(s) the plan meant to create do not exist under that name: ${names}${more}. It may have landed on a differently-named sheet instead.`;
  }

  if (renamedSheets.length > 0) {
    const sample = renamedSheets
      .slice(0, 3)
      .map((m) => `"${m.expected}" -> "${m.actual}"`)
      .join(', ');
    return `Applied, but ${renamedSheets.length} sheet(s) were created under a different name than requested: ${sample}${renamedSheets.length > 3 ? ' …' : ''}.`;
  }

  if (errors.length > 0) {
    const sample = errors.slice(0, 3).map((m) => `${m.sheet}!${m.cell} ${m.actual}`).join(', ');
    return `Applied, but ${errors.length} formula cell(s) returned an error: ${sample}${errors.length > 3 ? ' …' : ''}.`;
  }

  const sample = result.mismatches
    .slice(0, 3)
    .map((m) => `${m.sheet}!${m.cell}`)
    .join(', ');
  return `Applied, but ${result.mismatches.length} cell(s) do not match what was proposed: ${sample}${result.mismatches.length > 3 ? ' …' : ''}.`;
}
