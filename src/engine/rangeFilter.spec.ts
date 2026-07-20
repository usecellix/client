import { describe, expect, it } from 'vitest';
import { applyFilterOperator, filterDataRows } from './rangeFilter';
import { partitionActions } from './actionNormalizer';

describe('rangeFilter', () => {
  it('filters pending payment rows', () => {
    const { filteredRows } = filterDataRows(
      [
        ['Vendor', 'Payment Status'],
        ['A', 'Pending'],
        ['B', 'Paid'],
      ],
      true,
      { column: 'Payment Status', operator: 'equals', value: 'Pending' },
    );
    expect(filteredRows).toEqual([['A', 'Pending']]);
    expect(applyFilterOperator('Pending', { column: 'x', operator: 'equals', value: 'pending' })).toBe(
      true,
    );
  });

  it('supports length and regex operators for GSTIN-style checks', () => {
    expect(
      applyFilterOperator('SHORT', { column: 'GSTIN', operator: 'lengthNotEquals', value: 15 }),
    ).toBe(true);
    expect(
      applyFilterOperator('BAD!!', {
        column: 'GSTIN',
        operator: 'notMatchesRegex',
        value: '^[A-Za-z0-9]{15}$',
      }),
    ).toBe(true);
  });
});

describe('actionNormalizer native range actions', () => {
  it('routes COPY_FILTERED_RANGE to the rich engine', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'COPY_FILTERED_RANGE',
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:B4',
        hasHeaders: true,
        destSheet: 'Pending Payments',
        destStartCell: 'A1',
        mode: 'copy',
        filter: {
          column: 'Payment Status',
          operator: 'equals',
          value: 'Pending',
        },
      },
    ]);
    expect(unsupported).toHaveLength(0);
    expect(rich[0]).toEqual(
      expect.objectContaining({
        type: 'COPY_FILTERED_RANGE',
        destSheet: 'Pending Payments',
        mode: 'copy',
      }),
    );
  });

  it('routes MOVE_RANGE to the rich engine', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'MOVE_RANGE',
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:D10',
        destSheet: 'Archive',
        destStartCell: 'A1',
      },
    ]);
    expect(unsupported).toHaveLength(0);
    expect(rich[0]?.type).toBe('MOVE_RANGE');
  });

  it('routes FORMAT_MATCHING_ROWS to the rich engine', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'FORMAT_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A1:L51',
        hasHeaders: true,
        filter: { column: 'Payment Status', operator: 'equals', value: 'Pending' },
        format: { fillColor: '#FFC7CE' },
      },
    ]);
    expect(unsupported).toHaveLength(0);
    expect(rich[0]).toEqual(
      expect.objectContaining({
        type: 'FORMAT_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        filter: expect.objectContaining({ column: 'Payment Status', value: 'Pending' }),
        format: expect.objectContaining({ fillColor: '#FFC7CE' }),
      }),
    );
  });
});
