import { describe, expect, it } from 'vitest';
import { partitionActions } from './actionNormalizer';

describe('actionNormalizer', () => {
  it('routes DELETE_SHEET to the rich engine', () => {
    const { rich, unsupported } = partitionActions([
      { type: 'DELETE_SHEET', sheetName: 'Cellix' },
    ]);
    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Cellix' }]);
  });

  it('routes CREATE_SHEET to ADD_SHEET in the rich engine', () => {
    const { rich } = partitionActions([{ type: 'CREATE_SHEET', sheetName: 'Summary' }]);
    expect(rich).toEqual([{ type: 'ADD_SHEET', name: 'Summary' }]);
  });

  it('normalizes legacy CREATE_TABLE name and defaults headers to true', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CREATE_TABLE',
        sheetName: 'Purchase Register',
        range: 'A1:L51',
        name: 'PurchaseTable',
      },
    ]);

    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      {
        type: 'CREATE_TABLE',
        sheetName: 'Purchase Register',
        range: 'A1:L51',
        tableName: 'PurchaseTable',
        hasHeaders: true,
      },
    ]);
  });

  it('normalizes CREATE_CHART with safe placement defaults', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Purchase Register',
        sourceRange: 'A1:B10',
        chartType: 'ColumnClustered',
        title: 'Purchases by Month',
      },
    ]);

    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Purchase Register',
        sourceRange: 'A1:B10',
        chartType: 'ColumnClustered',
        title: 'Purchases by Month',
        startCell: 'A1',
        endCell: 'H16',
      },
    ]);
  });
});
