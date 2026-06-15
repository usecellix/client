import { describe, expect, it } from 'vitest';
import { partitionActions } from './actionNormalizer';

describe('actionNormalizer', () => {
  it('routes DELETE_SHEET to the rich engine', () => {
    const { rich, legacy } = partitionActions([
      { type: 'DELETE_SHEET', sheetName: 'Cellix' },
    ]);
    expect(legacy).toHaveLength(0);
    expect(rich).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Cellix' }]);
  });

  it('routes CREATE_SHEET to ADD_SHEET in the rich engine', () => {
    const { rich } = partitionActions([{ type: 'CREATE_SHEET', sheetName: 'Summary' }]);
    expect(rich).toEqual([{ type: 'ADD_SHEET', name: 'Summary' }]);
  });
});
