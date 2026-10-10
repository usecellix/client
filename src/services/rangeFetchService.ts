/* global Excel */

export interface RangeFetchResult {
  values: unknown[][];
}

function columnIndexToLetter(index: number): string {
  let n = index;
  let result = '';
  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

export async function fetchRangeData(
  sheetName: string,
  range: string,
): Promise<RangeFetchResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = sheet.getRange(range);
    targetRange.load('values');
    await context.sync();

    return {
      values: (targetRange.values ?? []) as unknown[][],
    };
  });
}

export async function navigateToCell(
  sheetName: string,
  row: number,
  col: number,
  /** When given, selects from `col` to this 0-based column on the same row. */
  endCol?: number,
): Promise<void> {
  return Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.activate();
    const start = `${columnIndexToLetter(col)}${row + 1}`;
    const address =
      endCol !== undefined && endCol > col ? `${start}:${columnIndexToLetter(endCol)}${row + 1}` : start;
    sheet.getRange(address).select();
    await context.sync();
  });
}
