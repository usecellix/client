import { SheetContext } from './context.types';

const MAX_ROWS_FULL = 80;
const MAX_ROWS_HEAD = 50;
const MAX_ROWS_TAIL = 10;
const MAX_SHEETS_FULL = 5;

export function compressSheetForPrompt(
  sheet: SheetContext,
  priority: 'active' | 'secondary',
): string {
  const lines: string[] = [];
  lines.push(
    `Sheet: "${sheet.name}" | ${sheet.rowCount}x${sheet.columnCount} | type: ${sheet.structure}`,
  );
  lines.push(`Range: ${sheet.usedRange}`);

  if (sheet.headers.length) {
    lines.push(`Headers: ${sheet.headers.slice(0, 20).join(' | ')}`);
  }

  if (sheet.formulaSummary) {
    lines.push(`Formulas: ${sheet.formulaSummary}`);
  }

  if (priority === 'secondary' && sheet.rowCount > 30) {
    lines.push(`Data preview (first 5 rows): ${JSON.stringify(sheet.values.slice(0, 5))}`);
    lines.push(`... [${sheet.rowCount - 10} rows omitted] ...`);
    lines.push(`Data preview (last 5 rows): ${JSON.stringify(sheet.values.slice(-5))}`);
    return lines.join('\n');
  }

  if (sheet.rowCount <= MAX_ROWS_FULL) {
    lines.push(`Values: ${JSON.stringify(sheet.values)}`);
    lines.push(`Formulas: ${JSON.stringify(sheet.formulas)}`);
  } else {
    lines.push(
      `Values (first ${MAX_ROWS_HEAD} rows): ${JSON.stringify(sheet.values.slice(0, MAX_ROWS_HEAD))}`,
    );
    lines.push(`... [${sheet.rowCount - MAX_ROWS_HEAD - MAX_ROWS_TAIL} rows omitted] ...`);
    lines.push(
      `Values (last ${MAX_ROWS_TAIL} rows): ${JSON.stringify(sheet.values.slice(-MAX_ROWS_TAIL))}`,
    );
  }

  return lines.join('\n');
}

export function buildPromptContext(
  activeSheetName: string,
  sheets: SheetContext[],
  namedRanges: { name: string; formula: string }[],
  tables: { name: string; sheetName: string; columnNames: string[] }[],
  selectedRange?: string,
): string {
  const parts: string[] = [];

  parts.push('=== WORKBOOK CONTEXT ===');
  parts.push(`Active sheet: ${activeSheetName}`);
  if (selectedRange) parts.push(`Selected range: ${selectedRange}`);
  parts.push(`Total sheets: ${sheets.length}`);

  if (namedRanges.length) {
    parts.push(
      `\nNamed ranges:\n${namedRanges.map((n) => `  ${n.name} = ${n.formula}`).join('\n')}`,
    );
  }

  if (tables.length) {
    parts.push(
      `\nTables:\n${tables.map((t) => `  ${t.name} (${t.sheetName}): ${t.columnNames.join(', ')}`).join('\n')}`,
    );
  }

  const activeSheet = sheets.find((s) => s.name === activeSheetName);
  if (activeSheet) {
    parts.push('\n--- Active Sheet ---');
    parts.push(compressSheetForPrompt(activeSheet, 'active'));
  }

  const secondarySheets = sheets
    .filter((s) => s.name !== activeSheetName && !s.isHidden)
    .slice(0, MAX_SHEETS_FULL - 1);

  for (const sheet of secondarySheets) {
    parts.push(`\n--- Sheet: ${sheet.name} ---`);
    parts.push(compressSheetForPrompt(sheet, 'secondary'));
  }

  const hiddenCount = sheets.filter((s) => s.isHidden).length;
  if (hiddenCount) {
    parts.push(`\n[${hiddenCount} hidden sheet(s) omitted]`);
  }

  return parts.join('\n');
}
