import { describe, expect, it, vi, beforeEach } from 'vitest';
import { jumpToWorkbookSource } from '@/utils/jumpToWorkbookSource';

describe('jumpToWorkbookSource', () => {
  beforeEach(() => {
    vi.stubGlobal('Excel', {
      run: vi.fn(async (fn: (ctx: unknown) => Promise<void>) => {
        const range = { select: vi.fn() };
        const sheet = { getRange: vi.fn(() => range) };
        const ctx = {
          workbook: {
            worksheets: {
              getItem: vi.fn(() => sheet),
              getActiveWorksheet: vi.fn(() => sheet),
            },
          },
          sync: vi.fn(),
        };
        await fn(ctx);
        return { range, sheet, ctx };
      }),
    });
  });

  it('selects the precedent range via Office.js for workbook refs', async () => {
    await jumpToWorkbookSource({
      documentType: 'workbook',
      documentId: 'wb',
      rowOrLine: 'Sheet2!C4:C40',
    });

    expect(Excel.run).toHaveBeenCalledTimes(1);
    const runMock = Excel.run as unknown as ReturnType<typeof vi.fn>;
    const ctxArg = runMock.mock.calls[0][0];
    const sheet = {
      getRange: vi.fn(() => ({ select: vi.fn() })),
    };
    const ctx = {
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet),
          getActiveWorksheet: vi.fn(() => sheet),
        },
      },
      sync: vi.fn(),
    };
    await ctxArg(ctx);
    expect(ctx.workbook.worksheets.getItem).toHaveBeenCalledWith('Sheet2');
    expect(sheet.getRange).toHaveBeenCalledWith('C4:C40');
  });

  it('rejects non-workbook document types', async () => {
    await expect(
      jumpToWorkbookSource({
        documentType: 'gstr2b',
        documentId: 'doc-1',
        rowOrLine: 12,
      }),
    ).rejects.toThrow(/Cannot jump in-sheet/);
  });
});
