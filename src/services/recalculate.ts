/* global Excel */

/**
 * Force Excel to recalculate before anything reads results back — TASKS.md #172.
 *
 * Cellix had no `calculate()` call anywhere. That is fine while a batch only
 * writes values, and wrong the moment a batch writes a cross-sheet formula
 * alongside the sheet it references: the formula can be evaluated against the
 * workbook as it stood mid-apply, and the stale result is what a read-back
 * then sees. Shortcut's own build log reached the same conclusion from the
 * other direction — it deferred adding its chart until after
 * `workbook.calculate()` so the series had real values to plot.
 *
 * Called between apply and verification, so #150's read-back judges settled
 * values rather than a half-calculated workbook. Without it a formula that is
 * perfectly correct can be reported as a failure, which is worse than no check
 * at all: it teaches the user to distrust a check that is right most of the
 * time.
 *
 * Never throws. A recalculation that fails is not a reason to turn a completed
 * write into an error — the same rule `verifyAppliedOutcomeSafe` follows.
 */
export async function recalculateWorkbookSafe(): Promise<boolean> {
  if (typeof Excel === 'undefined' || typeof Excel.run !== 'function') return false;

  try {
    await Excel.run(async (ctx) => {
      // 'Full' rather than the default recalculation: a formula that already
      // resolved to #REF! against a not-yet-created sheet is not "dirty" in
      // Excel's dependency tracking, so an ordinary recalculation can leave it
      // exactly as it was. Full is the one that reconsiders every cell.
      ctx.workbook.application.calculate(Excel.CalculationType.full);
      await ctx.sync();
    });
    return true;
  } catch (error) {
    console.warn('[Cellix] Workbook recalculation failed; verifying as-is:', error);
    return false;
  }
}
