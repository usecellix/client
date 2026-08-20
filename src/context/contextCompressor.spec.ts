import { describe, expect, it } from 'vitest';
import { buildPromptContext } from '@/context/contextCompressor';
import { SheetContext } from '@/context/context.types';

function makeSheet(values: unknown[][], rowCount = values.length): SheetContext {
  return {
    name: 'Sheet1',
    usedRange: 'A1:D10',
    rowCount,
    columnCount: 4,
    values,
    formulas: values.map(() => ['', '', '', '']),
    numberFormats: values.map(() => ['', '', '', '']),
    structure: 'data_table',
    headers: ['Item', 'Qty', 'Rate', 'Amount'],
    headerRowIndex: 0,
    formulaSummary: '',
    isHidden: false,
  };
}

describe('buildPromptContext TOON compression', () => {
  it('adds TOON format markers for larger tabular data', () => {
    const values: unknown[][] = [
      ['Item', 'Qty', 'Rate', 'Amount'],
      ['A', 1, 10, 10],
      ['B', 2, 10, 20],
      ['C', 3, 10, 30],
    ];
    const prompt = buildPromptContext('Sheet1', [makeSheet(values)], [], []);
    expect(prompt).toContain('sheetDataFormat: TOON');
    expect(prompt).toContain('sheetDataIncludedRows: 4/4');
  });

  it('keeps JSON for small/non-beneficial datasets', () => {
    const values: unknown[][] = [
      ['Item', 'Qty'],
      ['A', 1],
    ];
    const prompt = buildPromptContext('Sheet1', [makeSheet(values, 2)], [], []);
    expect(prompt).toContain('sheetDataFormat: JSON');
  });
});

describe('buildPromptContext — conditional-format rules (TASKS.md #38)', () => {
  it('lists existing rules with id/sheet/range/kind/summary', () => {
    const prompt = buildPromptContext('Sheet1', [makeSheet([['a']], 1)], [], [], undefined, [
      { id: 'cf-1', sheetName: 'Sheet1', range: 'B2:B50', ruleKind: 'cellValue', summary: 'greaterThan 1000' },
    ]);
    expect(prompt).toContain('Conditional format rules');
    expect(prompt).toContain('[cf-1] Sheet1!B2:B50 (cellValue: greaterThan 1000)');
  });

  it('omits the section entirely when there are no rules', () => {
    const prompt = buildPromptContext('Sheet1', [makeSheet([['a']], 1)], [], []);
    expect(prompt).not.toContain('Conditional format rules');
  });
});
