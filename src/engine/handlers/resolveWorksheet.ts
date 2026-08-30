/**
 * Single source of truth for worksheet resolution.
 *
 * This file used to hold a second, near-identical copy of `resolveWorksheet`
 * — same active-sheet fallback, separately maintained, imported by
 * `overwriteGuard` while the handlers imported the other one. Two copies of a
 * fallback this load-bearing is a trap: fix one, miss the other. Re-export
 * instead so there is exactly one implementation. See TASKS.md #137.
 */
export { resolveWorksheet } from '../sheetResolve';
