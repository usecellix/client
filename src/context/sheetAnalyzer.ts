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

export function detectHeaders(values: unknown[][]): string[] {
  for (const row of values.slice(0, 5)) {
    const nonEmpty = row.filter((v) => v !== null && v !== '');
    if (nonEmpty.length === 0) continue;

    const stringCount = nonEmpty.filter(
      (v) => typeof v === 'string' && Number.isNaN(Number(v)),
    ).length;
    if (stringCount / nonEmpty.length > 0.6) {
      return row.map((v) => String(v ?? ''));
    }
  }
  return [];
}
