import { describe, expect, it } from 'vitest';
import { stripSheetPrefix, parseRangeAddress, isLocalRangeAddress } from './addressUtils';
import { compareSortValues, parseSortableValue } from './sortCompare';

describe('stripSheetPrefix', () => {
  it('removes sheet qualifier from range addresses', () => {
    expect(stripSheetPrefix("'Purchases'!A1:M339")).toBe('A1:M339');
    expect(parseRangeAddress("'Purchases'!A1:M339")).toEqual({
      row: 0,
      col: 0,
      rowCount: 339,
      colCount: 13,
    });
  });

  it('validates local range addresses', () => {
    expect(isLocalRangeAddress('A1:M339')).toBe(true);
    expect(isLocalRangeAddress("'Purchases'!A1:M339")).toBe(true);
    expect(isLocalRangeAddress('not-a-range')).toBe(false);
  });
});

describe('compareSortValues', () => {
  it('sorts CGST Dr amounts in ascending numeric order', () => {
    const values = ['8533.98 Dr', '450.09 Dr', '12088.13 Dr', '146.82 Dr'];
    expect([...values].sort(compareSortValues)).toEqual([
      '146.82 Dr',
      '450.09 Dr',
      '8533.98 Dr',
      '12088.13 Dr',
    ]);
  });

  it('parses accounting suffixes', () => {
    expect(parseSortableValue('148.50 Dr')).toBe(148.5);
  });

  it('parses DD-MMM-YY dates for sort', () => {
    expect(typeof parseSortableValue('20-Feb-25')).toBe('number');
    expect(compareSortValues('05-Jan-25', '20-Feb-25')).toBeLessThan(0);
  });
});
