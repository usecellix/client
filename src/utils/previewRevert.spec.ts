import { describe, expect, it } from 'vitest';
import {
  buildPreviewRejectActions,
  buildStructuralRevertActions,
  partitionPreviewActions,
} from './previewRevert';

describe('previewRevert', () => {
  it('partitions sheet create from cell writes', () => {
    const actions = [
      { type: 'ADD_SHEET' as const, name: 'Cellix' },
      { type: 'SET_CELL' as const, sheetName: 'Cellix', row: 0, col: 0, value: 'Hi' },
    ];
    const { structural, deferred } = partitionPreviewActions(actions);
    expect(structural).toHaveLength(1);
    expect(deferred).toHaveLength(1);
  });

  it('builds DELETE_SHEET inverse for ADD_SHEET', () => {
    const inverse = buildStructuralRevertActions([
      { type: 'ADD_SHEET', name: 'Cellix' },
    ]);
    expect(inverse).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Cellix' }]);
  });

  it('builds full reject list when structural preview was applied', () => {
    const actions = [
      { type: 'ADD_SHEET', name: 'Cellix' },
      { type: 'SET_CELL', sheetName: 'Cellix', row: 0, col: 0, value: 'Hi' },
    ];
    const revert = buildPreviewRejectActions(actions, [], { structuralApplied: true });
    expect(revert).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Cellix' }]);
  });

  it('defers DELETE_SHEET until accept (reject is a no-op on workbook)', () => {
    const actions = [{ type: 'DELETE_SHEET', sheetName: 'Cellix' }];
    const { structural, deferred } = partitionPreviewActions(actions);
    expect(structural).toHaveLength(0);
    expect(deferred).toHaveLength(1);
  });
});
