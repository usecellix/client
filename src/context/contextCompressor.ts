import { SheetContext } from './context.types';
import { estimateTokens, toToon } from '@/utils/toon-adapter.util';

const MAX_ROWS_FULL = 80;
const MAX_ROWS_HEAD = 50;
const MAX_ROWS_TAIL = 10;
const MAX_SHEETS_FULL = 5;
const MAX_SHEET_TOKENS = 6000;

type EncodedSheetData = {
  format: 'TOON' | 'JSON';
  payload: string;
  includedRows: number;
  totalRows: number;
  truncated: boolean;
};

function headersForSheet(sheet: SheetContext): string[] {
  if (sheet.headers.length > 0) return sheet.headers;
  const firstRow = sheet.values[0] ?? [];
  return firstRow.map((value, index) => {
    const cell = String(value ?? '').trim();
    return cell || `Column_${index + 1}`;
  });
}

function rowsToObjects(headers: string[], rows: unknown[][]): Record<string, unknown>[] {
  return rows.map((row) => {
    const record: Record<string, unknown> = {};
    headers.forEach((header, index) => {
      record[header] = row[index] ?? '';
    });
    return record;
  });
}

function encodeSheetData(rows: Record<string, unknown>[]): EncodedSheetData {
  const totalRows = rows.length;
  if (totalRows === 0) {
    return {
      format: 'JSON',
      payload: '[]',
      includedRows: 0,
      totalRows: 0,
      truncated: false,
    };
  }

  const jsonPayload = JSON.stringify(rows);
  let effectiveRows = rows;
  let payload = toToon(effectiveRows);
  let tokenCount = estimateTokens(payload);

  while (tokenCount > MAX_SHEET_TOKENS && effectiveRows.length > 1) {
    effectiveRows = effectiveRows.slice(0, Math.max(1, Math.floor(effectiveRows.length * 0.8)));
    payload = toToon(effectiveRows);
    tokenCount = estimateTokens(payload);
  }

  const format = payload === JSON.stringify(effectiveRows) ? 'JSON' : 'TOON';
  const jsonTokens = estimateTokens(jsonPayload);

  if ((import.meta as { env?: { DEV?: boolean } }).env?.DEV) {
    // eslint-disable-next-line no-console
    console.debug('[Cellix] TOON compression ratio', {
      jsonTokens,
      compressedTokens: tokenCount,
      savedTokens: jsonTokens - tokenCount,
      savedPct: jsonTokens > 0 ? `${(((jsonTokens - tokenCount) / jsonTokens) * 100).toFixed(1)}%` : '0%',
      format,
    });
  }

  return {
    format,
    payload,
    includedRows: effectiveRows.length,
    totalRows,
    truncated: effectiveRows.length < totalRows,
  };
}

function appendSheetData(
  lines: string[],
  sheet: SheetContext,
  rows: unknown[][],
  label = 'sheetData',
): void {
  const headers = headersForSheet(sheet);
  const rowObjects = rowsToObjects(headers, rows);
  const encoded = encodeSheetData(rowObjects);

  lines.push(`${label}Format: ${encoded.format}`);
  lines.push(`${label}IncludedRows: ${encoded.includedRows}/${encoded.totalRows}`);
  if (encoded.truncated) {
    lines.push(`${label}Truncated: true (token budget ${MAX_SHEET_TOKENS})`);
  }
  lines.push(`${label}: ${encoded.payload}`);
}

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
    const previewRows = sheet.values.slice(0, 5);
    appendSheetData(lines, sheet, previewRows, 'sheetDataHead');
    lines.push(`... [${sheet.rowCount - 10} rows omitted] ...`);
    appendSheetData(lines, sheet, sheet.values.slice(-5), 'sheetDataTail');
    return lines.join('\n');
  }

  if (sheet.rowCount <= MAX_ROWS_FULL) {
    appendSheetData(lines, sheet, sheet.values);
    lines.push(`Formulas: ${JSON.stringify(sheet.formulas)}`);
  } else {
    appendSheetData(lines, sheet, sheet.values.slice(0, MAX_ROWS_HEAD), 'sheetDataHead');
    lines.push(`... [${sheet.rowCount - MAX_ROWS_HEAD - MAX_ROWS_TAIL} rows omitted] ...`);
    appendSheetData(lines, sheet, sheet.values.slice(-MAX_ROWS_TAIL), 'sheetDataTail');
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
