import { SortRangeAction } from '@/action.types';
import { isLocalRangeAddress, parseRangeAddress, stripSheetPrefix } from '../addressUtils';
import { resolveWorksheet } from '../sheetResolve';
import { compareSortValues } from '../sortCompare';
import { preserveNumberFormatsAroundWrite } from '@/services/formatGuard';

/* global Excel */

export async function handleSortRange(
  action: SortRangeAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  let rangeAddress = stripSheetPrefix(action.range);

  if (!parseRangeAddress(rangeAddress)) {
    const used = sheet.getUsedRange();
    if (!used) {
      throw new Error(`No data to sort on sheet "${action.sheetName}"`);
    }
    used.load('address');
    await ctx.sync();
    rangeAddress = stripSheetPrefix(used.address ?? '');
  }

  if (!isLocalRangeAddress(rangeAddress)) {
    throw new Error(`Invalid sort range "${action.range}"`);
  }

  const range = sheet.getRange(rangeAddress);
  range.load(['values', 'rowCount', 'columnCount']);
  await ctx.sync();

  const values = (range.values ?? []) as unknown[][];
  if (values.length < 2) return;

  const hasHeaders = action.hasHeaders ?? true;
  const headerRow = hasHeaders ? values[0] : null;
  const dataRows = (hasHeaders ? values.slice(1) : values).map((row) => [...row]);
  const key = action.key ?? 0;
  const ascending = action.ascending ?? true;

  // Sort order as data-row indices (0-based, relative to dataRows), applied
  // identically to values and to the numberFormat snapshot below — a row's
  // format always travels with the row it belongs to.
  const order = dataRows
    .map((_, index) => index)
    .sort((a, b) => {
      const cmp = compareSortValues(dataRows[a]![key], dataRows[b]![key]);
      return ascending ? cmp : -cmp;
    });

  const sortedValues = headerRow
    ? [headerRow, ...order.map((i) => dataRows[i]!)]
    : order.map((i) => dataRows[i]!);

  // Writing the reordered values via `range.values = ...` can make Excel's own
  // smart-entry parsing reinterpret a cell (e.g. re-detect a date and apply a
  // locale default format) if the explicit numberFormat isn't re-asserted
  // *after* that write completes, not merely alongside it in the same batch —
  // this is the exact "11/09/2025" -> "11092025" class of bug. Delegate to the
  // shared helper (built with sort specifically in mind, per its own
  // docstring) instead of setting values/numberFormat together by hand.
  await preserveNumberFormatsAroundWrite(
    range,
    ctx,
    () => {
      range.values = sortedValues;
    },
    (formats) => {
      const headerFormats = hasHeaders ? formats[0] : undefined;
      const dataFormats = hasHeaders ? formats.slice(1) : formats;
      const reordered = order.map((i) => dataFormats[i]!);
      return headerFormats ? [headerFormats, ...reordered] : reordered;
    },
  );
  await ctx.sync();
}
