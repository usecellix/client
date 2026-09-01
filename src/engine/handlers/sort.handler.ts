import { SortRangeAction } from '@/action.types';
import {
  columnIndexToLetter,
  isLocalRangeAddress,
  parseRangeAddress,
  stripSheetPrefix,
} from '../addressUtils';
import { resolveWorksheet } from '../sheetResolve';
import { compareSortValues } from '../sortCompare';
import { preserveNumberFormatsAroundWrite } from '@/services/formatGuard';
import { CellChange } from '@/types/changeSet';

/* global Excel */

/**
 * The real before/after cell diff for this sort, computed from the values
 * actually read off Excel — not the backend's shadow-workbook simulation,
 * which deliberately skips sparse ranges (mixed blank/filled cells) and so
 * can report "0 cells changed" for a sort that genuinely reordered the
 * sheet. Returned so the caller can report it to `/audit/apply` and give
 * Revert something real to work with (TASKS.md #93).
 */
function diffSortedGrid(
  sheetName: string,
  rangeAddress: string,
  before: unknown[][],
  after: unknown[][],
): CellChange[] {
  const bounds = parseRangeAddress(rangeAddress);
  if (!bounds) return [];

  const changes: CellChange[] = [];
  for (let r = 0; r < before.length; r += 1) {
    const beforeRow = before[r] ?? [];
    const afterRow = after[r] ?? [];
    for (let c = 0; c < beforeRow.length; c += 1) {
      const beforeVal = beforeRow[c] ?? null;
      const afterVal = afterRow[c] ?? null;
      if (String(beforeVal ?? '') === String(afterVal ?? '')) continue;
      changes.push({
        cell: `${columnIndexToLetter(bounds.col + c)}${bounds.row + r + 1}`,
        sheet: sheetName,
        before: beforeVal,
        after: afterVal,
        isHardcoded: true,
      });
    }
  }
  return changes;
}

export async function handleSortRange(
  action: SortRangeAction,
  ctx: Excel.RequestContext,
): Promise<{ sortedRangeChanges?: CellChange[] } | void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  let rangeAddress = stripSheetPrefix(action.range);

  const expandToUsedRange = async (): Promise<void> => {
    const used = sheet.getUsedRange();
    if (!used) {
      throw new Error(`No data to sort on sheet "${action.sheetName}"`);
    }
    used.load('address');
    await ctx.sync();
    rangeAddress = stripSheetPrefix(used.address ?? '');
  };

  if (!parseRangeAddress(rangeAddress)) {
    await expandToUsedRange();
  }

  if (!isLocalRangeAddress(rangeAddress)) {
    throw new Error(`Invalid sort range "${action.range}"`);
  }

  let range = sheet.getRange(rangeAddress);
  range.load(['values', 'rowCount', 'columnCount']);
  await ctx.sync();

  let values = (range.values ?? []) as unknown[][];

  // A parsed-but-too-small target (e.g. a stale single-cell selection like
  // "K13" left over from before the request) is not a legitimate sort range —
  // sorting fewer than 2 rows is a no-op. Silently returning here previously
  // reported false success ("1 Direct Change Applied") with the sheet never
  // touched. Expand to the sheet's used range once, same fallback already
  // used for an unparseable address, before concluding there's really
  // nothing to sort.
  if (values.length < 2) {
    await expandToUsedRange();
    range = sheet.getRange(rangeAddress);
    range.load(['values', 'rowCount', 'columnCount']);
    await ctx.sync();
    values = (range.values ?? []) as unknown[][];
  }

  if (values.length < 2) {
    throw new Error(`No data to sort on sheet "${action.sheetName}"`);
  }

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

  const changes = diffSortedGrid(action.sheetName, rangeAddress, values, sortedValues);
  return changes.length > 0 ? { sortedRangeChanges: changes } : undefined;
}
