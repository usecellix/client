/* global Excel */

/** Resolve a worksheet by name, falling back to the active sheet. */
export function resolveWorksheet(
  ctx: Excel.RequestContext,
  sheetName: string | undefined | null,
): Excel.Worksheet {
  const name = String(sheetName ?? '').trim();
  if (name.length > 0) {
    return ctx.workbook.worksheets.getItem(name);
  }
  return ctx.workbook.worksheets.getActiveWorksheet();
}
