import { describe, expect, it } from 'vitest';
import { countColumnFormats } from './contextAdapter';

describe('countColumnFormats (P1.5)', () => {
  const formats = [
    ['', ''],
    ['General', '#,##0.00 "Cr"'],
    ['General', '#,##0.00 "Cr"'],
    ['@', '#,##0.00 "Dr"'],
    ['', '#,##0.00 "Dr"'],
  ];
  it('skips the header row and General/@/blank, counts per format', () => {
    expect(countColumnFormats(formats, 0, 0)).toEqual([]);
    expect(countColumnFormats(formats, 0, 1)).toEqual([
      { format: '#,##0.00 "Cr"', count: 2 },
      { format: '#,##0.00 "Dr"', count: 2 },
    ]);
  });
  it('ignores formats above the header row', () => {
    expect(countColumnFormats([['0.0'], ['0.0'], ['0.0']], 1, 0)).toEqual([{ format: '0.0', count: 1 }]);
  });
});
