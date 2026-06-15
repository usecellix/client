import { SortRangeAction } from '@/action.types';
import { isLocalRangeAddress, parseRangeAddress, stripSheetPrefix } from '../addressUtils';
import { compareSortValues } from '../sortCompare';

/* global Excel */

export async function handleSortRange(
  action: SortRangeAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
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

  dataRows.sort((rowA, rowB) => {
    const cmp = compareSortValues(rowA[key], rowB[key]);
    return ascending ? cmp : -cmp;
  });

  range.values = headerRow ? [headerRow, ...dataRows] : dataRows;
  await ctx.sync();
}
