import { describe, expect, it } from 'vitest';
import { convertLegacyToRich } from './legacyConverter';
import { partitionActions } from './actionNormalizer';
import { SheetAction } from '@/types/sheet-actions';

describe('convertLegacyToRich', () => {
  it('converts SET_CELL row/col to address-based', () => {
    const action: SheetAction = {
      type: 'SET_CELL',
      sheetName: 'Sheet1',
      row: 0,
      col: 2,
      value: 'Total',
    };
    const rich = convertLegacyToRich(action);
    expect(rich).toEqual({
      type: 'SET_CELL',
      sheetName: 'Sheet1',
      address: 'C1',
      value: 'Total',
    });
  });

  it('converts ADD_ROW data array to APPEND_ROW', () => {
    const action: SheetAction = {
      type: 'ADD_ROW',
      sheetName: 'Sheet1',
      data: ['GST', '', '=C10*0.1'],
    };
    const rich = convertLegacyToRich(action);
    expect(rich).toEqual({
      type: 'APPEND_ROW',
      sheetName: 'Sheet1',
      values: ['GST', '', '=C10*0.1'],
    });
  });

  it('converts DELETE_ROW 0-based to 1-based rows', () => {
    const action: SheetAction = {
      type: 'DELETE_ROW',
      sheetName: 'Sheet1',
      row: 2,
    };
    const rich = convertLegacyToRich(action);
    expect(rich).toEqual({
      type: 'DELETE_ROW',
      sheetName: 'Sheet1',
      rows: [3],
    });
  });

  it('converts WRITE_TABLE to rich shape', () => {
    const action: SheetAction = {
      type: 'WRITE_TABLE',
      sheetName: 'Sheet1',
      headers: ['A', 'B'],
      rows: [[1, 2]],
    };
    const rich = convertLegacyToRich(action);
    expect(rich).toEqual({
      type: 'WRITE_TABLE',
      sheetName: 'Sheet1',
      headers: ['A', 'B'],
      rows: [[1, 2]],
    });
  });

  it('converts FORMAT_RANGE to range string', () => {
    const action: SheetAction = {
      type: 'FORMAT_RANGE',
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      rowCount: 1,
      colCount: 3,
      format: { bold: true },
    };
    const rich = convertLegacyToRich(action);
    expect(rich).toMatchObject({
      type: 'FORMAT_RANGE',
      sheetName: 'Sheet1',
      range: 'A1:C1',
    });
  });
});

describe('partitionActions', () => {
  it('routes executor-emitted legacy actions to rich engine', () => {
    const actions: SheetAction[] = [
      { type: 'SET_CELL', row: 1, col: 0, value: 'X', sheetName: 'Sheet1' },
      { type: 'ADD_ROW', data: ['a', 'b'], sheetName: 'Sheet1' },
      { type: 'CREATE_SHEET', sheetName: 'Summary' },
    ];

    const { rich, legacy } = partitionActions(actions);
    expect(rich).toHaveLength(3);
    expect(rich.map((a) => a.type)).toEqual(['SET_CELL', 'APPEND_ROW', 'ADD_SHEET']);
    expect(legacy).toHaveLength(0);
  });
});
