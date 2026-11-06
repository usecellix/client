import { describe, expect, it } from 'vitest';
import {
  applyFilterOperator,
  filterDataRows,
  findDuplicateOffsets,
  findMatchingRowOffsets,
  isTextNumberCell,
  parseDateMs,
  type RangeFilterSpec,
} from './rangeFilter';

const spec = (operator: RangeFilterSpec['operator'], value: string | number = '', extra: Partial<RangeFilterSpec> = {}): RangeFilterSpec => ({
  column: 'A',
  operator,
  value,
  ...extra,
});

describe('P1.7 rangeFilter operators', () => {
  describe('isBlank / isNotBlank', () => {
    it('empty, null and undefined are blank', () => {
      for (const v of ['', null, undefined]) expect(applyFilterOperator(v, spec('isBlank'))).toBe(true);
    });
    it('a whitespace-only cell is NOT blank by default', () => {
      expect(applyFilterOperator(' ', spec('isBlank'))).toBe(false);
      expect(applyFilterOperator(' ', spec('isNotBlank'))).toBe(true);
    });
    it('whitespace-only counts as blank with the trimSpaces flag', () => {
      expect(applyFilterOperator(' \u00A0 ', spec('isBlank', 'trimSpaces'))).toBe(true);
      expect(applyFilterOperator(' ', spec('isNotBlank', 'trimSpaces'))).toBe(false);
    });
    it('zero is not blank', () => {
      expect(applyFilterOperator(0, spec('isBlank'))).toBe(false);
    });
  });

  describe('isTextNumber', () => {
    it.each(['1,234', '5,000.50', '\u20B9 7,500', '(2,500)', ' 42 '])('text number: %j', (v) =>
      expect(applyFilterOperator(v, spec('isTextNumber'))).toBe(true));
    it.each([1234, 'abc', '', '01-04-2025', '20250401', null])('not a text number: %j', (v) =>
      expect(applyFilterOperator(v, spec('isTextNumber'))).toBe(false));
    it('helper agrees', () => expect(isTextNumberCell('1,234')).toBe(true));
  });

  describe('isInvalidGstin', () => {
    it('flags a checksum typo, short and wrong-state values', () => {
      for (const v of ['27AABCA1234F1Z7', '24DDDFR1111M', '50HHHPT5555S1ZC', '37CCCHD9876L1ZN'])
        expect(applyFilterOperator(v, spec('isInvalidGstin'))).toBe(true);
    });
    it('does not flag a valid GSTIN (any case, stray spaces)', () => {
      expect(applyFilterOperator('27AABCA1234F1Z9', spec('isInvalidGstin'))).toBe(false);
      expect(applyFilterOperator(' 27aabca1234f1z9 ', spec('isInvalidGstin'))).toBe(false);
    });
    it('does not flag a blank cell (that is isBlank)', () => {
      expect(applyFilterOperator('', spec('isInvalidGstin'))).toBe(false);
      expect(applyFilterOperator(null, spec('isInvalidGstin'))).toBe(false);
    });
  });

  describe('olderThan', () => {
    const asOn = '2026-04-30';
    it('text dates older than N days', () => {
      expect(applyFilterOperator('01-01-2026', spec('olderThan', 90, { asOn }))).toBe(true); // 119 days
      expect(applyFilterOperator('15-04-2026', spec('olderThan', 90, { asOn }))).toBe(false); // 15 days
    });
    it('exactly N days is not older than N', () => {
      expect(applyFilterOperator('2026-04-20', spec('olderThan', 10, { asOn }))).toBe(false);
      expect(applyFilterOperator('2026-04-19', spec('olderThan', 10, { asOn }))).toBe(true);
    });
    it('Excel serial dates', () => {
      // 46023 = 2026-01-01
      expect(parseDateMs(46023)).toBe(Date.UTC(2026, 0, 1));
      expect(applyFilterOperator(46023, spec('olderThan', 90, { asOn }))).toBe(true);
    });
    it('non-dates, blanks and plain small numbers never match', () => {
      for (const v of ['hello', '', null, 12]) expect(applyFilterOperator(v, spec('olderThan', 1, { asOn }))).toBe(false);
    });
    it('a two-digit year reads as 20yy', () => {
      expect(parseDateMs('1.4.25')).toBe(Date.UTC(2025, 3, 1));
    });
  });

  describe('regex (existing) stays safe', () => {
    it('an invalid pattern matches nothing instead of throwing', () => {
      expect(applyFilterOperator('abc', spec('matchesRegex', '([unclosed'))).toBe(false);
      expect(applyFilterOperator('abc', spec('notMatchesRegex', '([unclosed'))).toBe(false);
    });
  });

  describe('isDuplicate', () => {
    const rows = [
      ['Party', 'Invoice'],
      ['A', 'INV-1'],
      ['B', 'INV-2'],
      ['a ', 'inv-1'], // case/space-different duplicate of row 1
      ['C', ''],
      ['D', ''], // blank keys are never duplicates
      ['B', 'INV-2'],
    ];
    const data = rows.slice(1);

    it('keep first: later copies match; case and spaces ignored', () => {
      expect(findDuplicateOffsets(data, [1], 'first')).toEqual([2, 5]);
    });
    it('keep last: earlier copies match', () => {
      expect(findDuplicateOffsets(data, [1], 'last')).toEqual([0, 1]);
    });
    it('multi-column key', () => {
      expect(findDuplicateOffsets(data, [0, 1], 'first')).toEqual([2, 5]);
      expect(findDuplicateOffsets([['a', '1'], ['a', '2'], ['a', '1']], [0, 1])).toEqual([2]);
    });
    it('goes through findMatchingRowOffsets / filterDataRows with header offsets', () => {
      const filter = spec('isDuplicate', '', { column: 'Invoice' });
      expect(findMatchingRowOffsets(rows, true, filter)).toEqual([3, 6]);
      expect(filterDataRows(rows, true, filter).filteredRows).toEqual([['a ', 'inv-1'], ['B', 'INV-2']]);
    });
    it('keyColumns override the column', () => {
      const filter = spec('isDuplicate', '', { column: 'Invoice', keyColumns: ['Party', 'Invoice'], keep: 'first' });
      expect(findMatchingRowOffsets(rows, true, filter)).toEqual([3, 6]);
    });
  });

  it('still behaves for the classic operators', () => {
    expect(applyFilterOperator('Pending', spec('equals', 'pending'))).toBe(true);
    expect(applyFilterOperator(5, spec('greaterThan', 3))).toBe(true);
  });
});
