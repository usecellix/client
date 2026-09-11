import { HideGridlinesAction } from '@/action.types';
import { resolveWorksheet } from '../sheetResolve';

/* global Excel */

/**
 * HIDE_GRIDLINES — TASKS.md #169.
 *
 * A financial tracker reads as a built document rather than a spreadsheet when
 * the grid is off; every reference build we compared against turns it off, and
 * Cellix had no way to express it at all.
 *
 * `showGridlines` is a worksheet VIEW property, so it is per-sheet and must be
 * set on each one — a multi-sheet build emits one action per sheet rather than
 * a single workbook-wide switch.
 */
export async function handleHideGridlines(
  action: HideGridlinesAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  sheet.showGridlines = action.showGridlines === true;
  await ctx.sync();
}
