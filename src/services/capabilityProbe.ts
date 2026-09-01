/* global Excel */

/**
 * Excel capability probe — TASKS.md #152.
 *
 * Shortcut wrote `=VSTACK("a","b")` to a scratch cell to confirm dynamic-array
 * support *before* trusting it, and repaired its own `#REF!` when a formula
 * failed. Cellix currently ships #142's consolidation formula — built on
 * `LET`/`VSTACK`/`HSTACK`/`FILTER`/`BYROW`/`LAMBDA`/`EXPAND`/`DROP` — as an
 * untested assumption. On pre-365 Excel every one of those is `#NAME?`, and
 * nothing in the pipeline notices.
 *
 * There is no "evaluate this formula" API in Office.js, so a probe must write.
 * This writes ONE cell, in the far corner of the workbook's own scratch space,
 * reads the result, and clears it — and it is deliberately the *same function
 * set the consolidation formula actually uses*, so a pass means that specific
 * formula will work, not merely that "some dynamic arrays exist".
 */

export interface ExcelCapabilities {
  /** True when LET/VSTACK/HSTACK/FILTER/BYROW/LAMBDA/EXPAND/DROP all evaluate. */
  dynamicArrays: boolean;
  /** False when the probe could not run — capability is unknown, not absent. */
  probed: boolean;
  reason?: string;
}

export const UNKNOWN_CAPABILITIES: ExcelCapabilities = {
  dynamicArrays: false,
  probed: false,
  reason: 'not probed',
};

/**
 * A single-cell, non-spilling exercise of the exact function set
 * `consolidation-pass.util.ts` emits. Wrapping in COUNT keeps the result
 * scalar, so the probe never spills into a neighbouring cell.
 *
 * Evaluates to 2 on a supporting host; `#NAME?` on a host missing any one of
 * the functions.
 */
export const CAPABILITY_PROBE_FORMULA =
  '=COUNT(LET(rows,VSTACK(HSTACK(EXPAND(1,2,1,1),VSTACK(10,20))),' +
  'keep,BYROW(DROP(rows,,1),LAMBDA(r,COUNTA(r)>0)),FILTER(DROP(rows,,1),keep,"")))';

/**
 * Far corner of the sheet — chosen so a probe can never collide with user data
 * or with anything the current build is about to write. Cleared immediately.
 */
const PROBE_CELL = 'XFD1048575';

let cached: ExcelCapabilities | null = null;

/** Test seam — clears the per-session cache. */
export function resetCapabilityCache(): void {
  cached = null;
}

function isNameError(value: unknown): boolean {
  return typeof value === 'string' && /^#(NAME|VALUE|N\/A)[!?]?$/i.test(value.trim());
}

/**
 * Probe once per session and cache. Excel's function set cannot change
 * mid-session, so re-probing would be pure cost.
 *
 * Never throws: a probe failure yields `probed: false`, which callers must
 * treat as "unknown", never as "unsupported" — silently downgrading a capable
 * host to the fallback formula would be its own regression.
 */
export async function probeExcelCapabilities(): Promise<ExcelCapabilities> {
  if (cached) return cached;

  if (typeof Excel === 'undefined' || typeof Excel.run !== 'function') {
    cached = { ...UNKNOWN_CAPABILITIES, reason: 'Office.js unavailable' };
    return cached;
  }

  try {
    const result = await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(PROBE_CELL);

      cell.formulas = [[CAPABILITY_PROBE_FORMULA]];
      await ctx.sync();

      cell.load(['values', 'text']);
      await ctx.sync();

      const value = (cell.values as unknown[][] | undefined)?.[0]?.[0];

      // Always clear, on every path — the probe must leave no trace.
      cell.clear(Excel.ClearApplyTo.all);
      await ctx.sync();

      if (isNameError(value)) {
        return {
          dynamicArrays: false,
          probed: true,
          reason: `probe returned ${String(value)}`,
        } satisfies ExcelCapabilities;
      }

      return {
        dynamicArrays: typeof value === 'number' && value > 0,
        probed: true,
      } satisfies ExcelCapabilities;
    });

    cached = result;
    return cached;
  } catch (error) {
    const reason = error instanceof Error ? error.message : 'probe failed';
    console.warn('[Cellix] Excel capability probe could not run:', error);
    cached = { dynamicArrays: false, probed: false, reason };
    return cached;
  }
}
