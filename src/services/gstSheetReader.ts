/* global Excel */

/**
 * Full used-range reads for GST recon (samples in workbookContext are not enough).
 */

export interface SheetGrid {
  name: string;
  values: unknown[][];
  headers: string[];
  rowCount: number;
  columnCount: number;
}

export async function listWorksheetNames(): Promise<string[]> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();
    return sheets.items.map((s) => s.name);
  });
}

/** Peek header rows for all (or filtered) sheets without loading every data cell. */
export async function readAllSheetHeaders(
  sheetNames?: string[],
): Promise<Array<{ name: string; headers: string[] }>> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();

    const names = sheetNames?.length
      ? sheetNames
      : sheets.items.map((s) => s.name);

    const result: Array<{ name: string; headers: string[] }> = [];

    for (const name of names) {
      try {
        const sheet = ctx.workbook.worksheets.getItem(name);
        sheet.load('name');
        const used = sheet.getUsedRange();
        used.load(['rowCount', 'columnCount']);
        await ctx.sync();

        if (!used.rowCount || !used.columnCount) {
          result.push({ name: sheet.name, headers: [] });
          continue;
        }

        const headerRange = sheet.getRangeByIndexes(0, 0, 1, used.columnCount);
        headerRange.load('values');
        await ctx.sync();

        const headers = ((headerRange.values?.[0] as unknown[]) ?? []).map((c) =>
          String(c ?? '').trim(),
        );
        result.push({ name: sheet.name, headers });
      } catch {
        result.push({ name, headers: [] });
      }
    }
    return result;
  });
}

export async function readSheetsFull(sheetNames: string[]): Promise<SheetGrid[]> {
  if (!sheetNames.length) return [];

  return Excel.run(async (ctx) => {
    const results: SheetGrid[] = [];
    for (const name of sheetNames) {
      const sheet = ctx.workbook.worksheets.getItem(name);
      sheet.load('name');
      await ctx.sync();

      try {
        const used = sheet.getUsedRange();
        used.load(['values', 'rowCount', 'columnCount']);
        await ctx.sync();

        const values = ((used.values as unknown[][]) ?? []).map((row: unknown[]) => [
          ...(row ?? []),
        ]);
        const headers = (values[0] ?? []).map((c: unknown) => String(c ?? '').trim());
        results.push({
          name: sheet.name,
          values,
          headers,
          rowCount: used.rowCount,
          columnCount: used.columnCount,
        });
      } catch {
        results.push({
          name: sheet.name,
          values: [],
          headers: [],
          rowCount: 0,
          columnCount: 0,
        });
      }
    }
    return results;
  });
}
