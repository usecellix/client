import { describe, expect, it, vi, afterEach } from 'vitest';
import { navigateToCell } from './rangeFetchService';

/* global globalThis */

function installExcelMock() {
  const select = vi.fn();
  const activate = vi.fn();
  const getRange = vi.fn(() => ({ select }));
  const sheet = { activate, getRange };
  const getItem = vi.fn(() => sheet);
  const getActiveWorksheet = vi.fn(() => sheet);

  (globalThis as Record<string, unknown>).Excel = {
    run: async (cb: (ctx: unknown) => Promise<void>) =>
      cb({ workbook: { worksheets: { getItem, getActiveWorksheet } }, sync: vi.fn(async () => undefined) }),
  };
  return { select, activate, getRange, getItem, getActiveWorksheet };
}

describe('navigateToCell', () => {
  const originalExcel = (globalThis as Record<string, unknown>).Excel;
  afterEach(() => {
    (globalThis as Record<string, unknown>).Excel = originalExcel;
  });

  it('selects the whole row span when endCol is given', async () => {
    const { getRange, getItem, activate, select } = installExcelMock();
    await navigateToCell('Bank Statement', 444, 0, 8);

    expect(getItem).toHaveBeenCalledWith('Bank Statement');
    expect(activate).toHaveBeenCalledTimes(1);
    expect(getRange).toHaveBeenCalledWith('A445:I445');
    expect(select).toHaveBeenCalledTimes(1);
  });

  it('selects one cell when no endCol is given', async () => {
    const { getRange } = installExcelMock();
    await navigateToCell('Bank Statement', 4, 2);
    expect(getRange).toHaveBeenCalledWith('C5');
  });

  it('selects one cell when endCol does not go past col', async () => {
    const { getRange } = installExcelMock();
    await navigateToCell('Bank Statement', 4, 2, 2);
    expect(getRange).toHaveBeenCalledWith('C5');
  });

  it('uses the active sheet when no sheet name is given', async () => {
    const { getActiveWorksheet, getRange } = installExcelMock();
    await navigateToCell('', 0, 27, 28);
    expect(getActiveWorksheet).toHaveBeenCalled();
    expect(getRange).toHaveBeenCalledWith('AB1:AC1');
  });
});
