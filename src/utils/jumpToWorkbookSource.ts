import { SourceRef } from '@/types/changeSet';

/**
 * Select a workbook precedent range in the live Excel sheet (Office.js).
 * For non-workbook refs, the caller should open an external source viewer instead.
 */
export async function jumpToWorkbookSource(ref: SourceRef): Promise<void> {
  if (ref.documentType !== 'workbook') {
    throw new Error(`Cannot jump in-sheet for documentType=${ref.documentType}`);
  }

  const raw = String(ref.rowOrLine).trim();
  if (!raw) {
    throw new Error('Source ref has empty rowOrLine');
  }

  const bang = raw.lastIndexOf('!');
  const sheetName = bang >= 0 ? raw.slice(0, bang).replace(/^'|'$/g, '') : undefined;
  const address = bang >= 0 ? raw.slice(bang + 1) : raw;

  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.select();
    await context.sync();
  });
}
