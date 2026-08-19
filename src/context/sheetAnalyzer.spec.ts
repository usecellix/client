import { describe, expect, it } from 'vitest';
import { detectHeaders } from './sheetAnalyzer';

/**
 * Regression coverage: a lone title cell (e.g. a merged "ABC Corp — Purchase
 * Register" row) used to satisfy the bare text-ratio heuristic on its own and
 * get misidentified as the header row, before the scan ever reached the real
 * headers underneath it.
 */
describe('detectHeaders', () => {
  it('does not mistake a single-cell title row for headers', () => {
    const values = [
      ['ABC Corp — Purchase Register FY24'],
      [],
      ['Date', 'Supplier', 'Invoice No', 'Amount'],
      ['01-04-2024', 'Acme Ltd', 'INV-001', 12500],
    ];
    const result = detectHeaders(values);
    expect(result.headerRowIndex).toBe(2);
    expect(result.headers).toEqual(['Date', 'Supplier', 'Invoice No', 'Amount']);
  });

  it('still finds headers on row 0 for a normal table', () => {
    const values = [
      ['Name', 'Age', 'City'],
      ['Alice', 30, 'Mumbai'],
    ];
    const result = detectHeaders(values);
    expect(result.headerRowIndex).toBe(0);
    expect(result.headers).toEqual(['Name', 'Age', 'City']);
  });

  it('returns empty headers and index 0 when nothing qualifies', () => {
    const values = [[1, 2, 3], [4, 5, 6]];
    const result = detectHeaders(values);
    expect(result.headerRowIndex).toBe(0);
    expect(result.headers).toEqual([]);
  });
});
