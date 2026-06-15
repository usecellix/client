export interface FormulaSummary {
  totalFormulas: number;
  crossSheetRefs: string[];
  functionTypes: string[];
  aggregationRows: number[];
  humanReadable: string;
}

const FORMULA_FUNCTIONS = [
  'SUM',
  'SUMIF',
  'SUMIFS',
  'AVERAGE',
  'AVERAGEIF',
  'COUNT',
  'COUNTA',
  'COUNTIF',
  'COUNTIFS',
  'IF',
  'IFS',
  'AND',
  'OR',
  'NOT',
  'VLOOKUP',
  'HLOOKUP',
  'INDEX',
  'MATCH',
  'XLOOKUP',
  'IFERROR',
  'IFNA',
  'MAX',
  'MIN',
  'LARGE',
  'SMALL',
  'ROUND',
  'ROUNDUP',
  'ROUNDDOWN',
  'NPV',
  'IRR',
  'PMT',
  'FV',
  'PV',
  'TEXT',
  'LEFT',
  'RIGHT',
  'MID',
  'LEN',
  'CONCAT',
  'DATE',
  'YEAR',
  'MONTH',
  'DAY',
  'EDATE',
  'EOMONTH',
];

export function summarizeFormulas(formulas: string[][]): FormulaSummary {
  const allFormulas: { row: number; col: number; formula: string }[] = [];

  for (let r = 0; r < formulas.length; r += 1) {
    for (let c = 0; c < formulas[r].length; c += 1) {
      const f = formulas[r][c];
      if (typeof f === 'string' && f.startsWith('=')) {
        allFormulas.push({ row: r, col: c, formula: f });
      }
    }
  }

  const crossSheetRefs = [
    ...new Set(
      allFormulas
        .filter((f) => f.formula.includes('!'))
        .flatMap((f) => f.formula.match(/[A-Za-z0-9_\s]+![A-Z]+\d+/g) ?? []),
    ),
  ];

  const functionTypes = [
    ...new Set(
      allFormulas.flatMap((f) =>
        FORMULA_FUNCTIONS.filter((fn) => f.formula.toUpperCase().includes(`${fn}(`)),
      ),
    ),
  ];

  const aggregationRows: number[] = [];
  for (let r = 0; r < formulas.length; r += 1) {
    const rowFormulas = formulas[r].filter(
      (f) => typeof f === 'string' && f.includes('SUM('),
    );
    if (rowFormulas.length > formulas[r].length * 0.5) {
      aggregationRows.push(r + 1);
    }
  }

  const humanReadable = [
    `Total formulas: ${allFormulas.length}`,
    functionTypes.length ? `Functions used: ${functionTypes.join(', ')}` : '',
    crossSheetRefs.length
      ? `Cross-sheet refs: ${crossSheetRefs.slice(0, 5).join(', ')}${crossSheetRefs.length > 5 ? '...' : ''}`
      : '',
    aggregationRows.length
      ? `Aggregation rows (likely totals): ${aggregationRows.slice(0, 5).join(', ')}`
      : '',
  ]
    .filter(Boolean)
    .join('. ');

  return {
    totalFormulas: allFormulas.length,
    crossSheetRefs,
    functionTypes,
    aggregationRows,
    humanReadable,
  };
}
