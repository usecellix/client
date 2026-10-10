import { describe, expect, it } from 'vitest';
import { partitionActions } from './actionNormalizer';
import type { SheetAction } from '@/types/sheet-actions';

/**
 * TASKS.md #400 — Revert applies the server's inverse actions through the same
 * conversion as any other action. These are the exact shapes
 * `Server/src/audit/format-snapshot.ts` builds; if the conversion dropped
 * `clearFill`, the 'general' alignment or a column count, a revert would report
 * success and leave part of the formatting behind.
 */
describe('partitionActions — the inverse actions of a formatting revert', () => {
  it('keeps every restored value on a FORMAT_RANGE and addresses the same cells', () => {
    const inverse = [
      {
        type: 'FORMAT_RANGE',
        sheetName: 'Monthly Summary',
        row: 2,
        col: 0,
        rowCount: 1,
        colCount: 4,
        format: { bold: false, clearFill: true, fontColor: '#000000', horizontalAlignment: 'general', numberFormat: 'General' },
      },
    ] as unknown as SheetAction[];

    const { rich, unsupported } = partitionActions(inverse);

    expect(unsupported).toEqual([]);
    expect(rich).toEqual([
      {
        type: 'FORMAT_RANGE',
        sheetName: 'Monthly Summary',
        range: 'A3:D3',
        format: { bold: false, clearFill: true, fontColor: '#000000', horizontalAlignment: 'general', numberFormat: 'General' },
      },
    ]);
  });

  it('keeps the column, the count and the width on a SET_COLUMN_WIDTH', () => {
    const inverse = [
      { type: 'SET_COLUMN_WIDTH', sheetName: 'Monthly Summary', col: 0, colCount: 4, width: 64 },
    ] as unknown as SheetAction[];

    const { rich, unsupported } = partitionActions(inverse);

    expect(unsupported).toEqual([]);
    expect(rich).toEqual([{ type: 'SET_COLUMN_WIDTH', sheetName: 'Monthly Summary', col: 0, colCount: 4, width: 64 }]);
  });
});
