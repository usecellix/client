/* global Excel */

/** Resolve a worksheet by name, or the active sheet when name is missing/blank. */
export function resolveWorksheet(
  ctx: Excel.RequestContext,
  sheetName: string | undefined | null,
): Excel.Worksheet {
  const name = typeof sheetName === 'string' ? sheetName.trim() : '';
  if (name.length > 0) {
    return ctx.workbook.worksheets.getItem(name);
  }
  return ctx.workbook.worksheets.getActiveWorksheet();
}
