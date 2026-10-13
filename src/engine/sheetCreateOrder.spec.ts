import { describe, expect, it } from 'vitest';
import { hoistSheetCreates } from './sheetCreateOrder';

type A = { type: string } & Record<string, unknown>;
const order = (actions: A[]) => hoistSheetCreates(actions).map((a) => `${a.type}:${a.name ?? a.sheetName ?? ''}`);

/** TASKS.md #352 — a sheet is created before anything in the card touches it. */
describe('hoistSheetCreates', () => {
  it('moves the create ahead of a write ordered before it (the live Lists card)', () => {
    expect(
      order([
        { type: 'SET_CELL', sheetName: 'Lists', address: 'A1' },
        { type: 'ADD_SHEET', name: 'Lists' },
        { type: 'BATCH_SET', sheetName: 'Lists' },
        { type: 'SET_COLUMN_WIDTH', sheetName: 'Lists' },
      ]),
    ).toEqual(['ADD_SHEET:Lists', 'SET_CELL:Lists', 'BATCH_SET:Lists', 'SET_COLUMN_WIDTH:Lists']);
  });

  it('moves it only as far as the first use, leaving other sheets in place', () => {
    expect(
      order([
        { type: 'SET_CELL', sheetName: 'January' },
        { type: 'FORMAT_RANGE', sheetName: 'Main' },
        { type: 'ADD_SHEET', name: 'Main' },
      ]),
    ).toEqual(['SET_CELL:January', 'ADD_SHEET:Main', 'FORMAT_RANGE:Main']);
  });

  it('counts cross-sheet destinations as a use', () => {
    expect(
      order([
        { type: 'COPY_FILTERED_RANGE', sourceSheet: 'January', destSheet: 'Archive' },
        { type: 'ADD_SHEET', name: 'Archive' },
      ]),
    ).toEqual(['ADD_SHEET:Archive', 'COPY_FILTERED_RANGE:']);
  });

  it('keeps delete-then-recreate in that order', () => {
    const actions: A[] = [
      { type: 'DELETE_SHEET', sheetName: 'Main' },
      { type: 'ADD_SHEET', name: 'Main' },
    ];
    expect(hoistSheetCreates(actions)).toEqual(actions);
  });

  it('does not move a create past a rename into or out of that name', () => {
    const actions: A[] = [
      { type: 'SET_CELL', sheetName: 'Main' },
      { type: 'RENAME_SHEET', oldName: 'Main', newName: 'Old Main' },
      { type: 'ADD_SHEET', name: 'Main' },
    ];
    expect(hoistSheetCreates(actions)).toEqual(actions);
  });

  it('leaves a copy-based create alone — its source may be created in between', () => {
    const actions: A[] = [
      { type: 'SET_CELL', sheetName: 'February' },
      { type: 'ADD_SHEET', name: 'January' },
      { type: 'ADD_SHEET', name: 'February', copyFrom: 'January' },
    ];
    expect(hoistSheetCreates(actions)).toEqual(actions);
  });

  it('is a no-op when creates already come first', () => {
    const actions: A[] = [
      { type: 'ADD_SHEET', name: 'Main' },
      { type: 'BATCH_SET', sheetName: 'Main' },
    ];
    expect(hoistSheetCreates(actions)).toEqual(actions);
  });

  it('does not mutate its input', () => {
    const actions: A[] = [
      { type: 'SET_CELL', sheetName: 'Lists' },
      { type: 'ADD_SHEET', name: 'Lists' },
    ];
    const copy = [...actions];
    hoistSheetCreates(actions);
    expect(actions).toEqual(copy);
  });
});
