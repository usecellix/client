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

  it('defers INSERT_COLUMN and ADD_ROW until accept', () => {
    const actions = [
      {
        type: 'INSERT_COLUMN' as const,
        sheetName: 'Purchase Register',
        columnName: 'Net of Tax',
        position: 'afterLastColumn' as const,
        formula: '=J{row}-I{row}',
      },
      {
        type: 'ADD_ROW' as const,
        sheetName: 'Purchase Register',
        data: ['a', 'b'],
      },
    ];
    const { structural, deferred } = partitionPreviewActions(actions);
    expect(structural).toHaveLength(0);
    expect(deferred).toEqual(actions);
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

  it('defers SORT_RANGE until accept (reject is a no-op on workbook)', () => {
    const actions = [
      {
        type: 'SORT_RANGE' as const,
        sheetName: 'Purchase Register',
        range: 'A1:L51',
        key: 8,
        ascending: false,
        hasHeaders: true,
      },
    ];
    const { structural, deferred } = partitionPreviewActions(actions);
    expect(structural).toHaveLength(0);
    expect(deferred).toEqual(actions);

    // Nothing applied during preview → reject builds no revert actions.
    const revert = buildPreviewRejectActions(actions, [], {
      structuralApplied: false,
      deferredApplied: false,
    });
    expect(revert).toEqual([]);
  });

  it('defers COPY_FILTERED_RANGE and MOVE_RANGE until accept', () => {
    const actions = [
      {
        type: 'COPY_FILTERED_RANGE' as const,
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:L51',
        hasHeaders: true,
        destSheet: 'Pending Payments',
        destStartCell: 'A1',
        mode: 'copy' as const,
      },
      {
        type: 'MOVE_RANGE' as const,
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:D10',
        destSheet: 'Archive',
        destStartCell: 'A1',
      },
    ];
    const { structural, deferred } = partitionPreviewActions(actions);
    expect(structural).toHaveLength(0);
    expect(deferred).toEqual(actions);
  });

  it('defers CREATE_CHART, UPDATE_CHART, and AGGREGATE_TABLE until accept', () => {
    const actions = [
      {
        type: 'CREATE_CHART' as const,
        sheetName: 'Dashboard',
        sourceSheetName: 'Dashboard',
        sourceRange: 'A4:B9',
        chartType: 'ColumnClustered',
        chartId: 'Chart_1',
      },
      {
        type: 'UPDATE_CHART' as const,
        sheetName: 'Dashboard',
        chartId: 'Chart_1',
        chartType: 'BarClustered',
      },
      {
        type: 'AGGREGATE_TABLE' as const,
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:B10',
        groupByColumn: 'Supplier',
        aggregations: [{ column: 'Total Amount', fn: 'sum' as const, outputLabel: 'Spend' }],
        destSheet: 'Dashboard',
        destStartCell: 'A4',
      },
    ];
    const { structural, deferred } = partitionPreviewActions(actions);
    expect(structural).toHaveLength(0);
    expect(deferred).toEqual(actions);
  });
});
