import { SheetStructure } from './context.types';

export function detectSheetStructure(
  values: unknown[][],
  formulas: string[][],
): SheetStructure {
  if (!values.length) return 'unknown';

  const allCells = values.flat().filter((v) => v !== null && v !== '');
  const allFormulas = formulas
    .flat()
    .filter((f) => typeof f === 'string' && f.startsWith('='));

  if (allCells.length === 0) return 'unknown';

  const formulaRatio = allFormulas.length / allCells.length;

  const headerRow = values[0] ?? [];
  const hasYearHeaders = headerRow.some((h) => /^20\d{2}$/.test(String(h)));
  const hasFYHeaders = headerRow.some((h) => /FY\d{2,4}/i.test(String(h)));
  const hasQuarterHeaders = headerRow.some((h) => /Q[1-4]/i.test(String(h)));

  if (formulaRatio > 0.25 && (hasYearHeaders || hasFYHeaders || hasQuarterHeaders)) {
    return 'financial_model';
  }

  if (formulaRatio < 0.05) {
    return 'data_table';
  }

  return 'report';
}

export interface DetectedHeaders {
  headers: string[];
  /** 0-based row index within `values` the headers came from (0 when nothing matched). */
  headerRowIndex: number;
}

/**
 * A row only qualifies as headers if it is BOTH mostly text AND uses a
 * meaningful share of the sheet's width. Ratio alone misfires on a lone title
 * cell (e.g. a merged "ABC Corp — Purchase Register" cell) — one non-empty cell
 * that is 100% text satisfies a bare ratio check instantly, picking the title
 * row before the scan ever reaches the real headers underneath it.
 */
export function detectHeaders(values: unknown[][]): DetectedHeaders {
  const columnCount = values.reduce((max, row) => Math.max(max, row.length), 0);
  const minWidth = Math.max(2, Math.ceil(columnCount * 0.5));

  for (let rowIndex = 0; rowIndex < Math.min(values.length, 5); rowIndex += 1) {
    const row = values[rowIndex];
    const nonEmpty = row.filter((v) => v !== null && v !== '');
    if (nonEmpty.length === 0 || nonEmpty.length < minWidth) continue;

    const stringCount = nonEmpty.filter(
      (v) => typeof v === 'string' && Number.isNaN(Number(v)),
    ).length;
    if (stringCount / nonEmpty.length > 0.6) {
      return { headers: row.map((v) => String(v ?? '')), headerRowIndex: rowIndex };
    }
  }
  return { headers: [], headerRowIndex: 0 };
}
