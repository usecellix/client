import { describe, expect, it, vi } from 'vitest';
import { handleWorksheetAction } from './worksheet.handler';
import type { RichAction } from '@/action.types';

/**
 * Regression: creating month sheets then writing into them failed the whole
 * batch with Office.js "The requested resource doesn't exist."
 *
 * The engine tries this handler first for EVERY action. It used to call
 * resolveWorksheet() — i.e. `worksheets.getItem(name)` — before checking whether
 * it handles the type at all. For ADD_SHEET/CREATE_SHEET the target sheet does
 * not exist yet (it is about to be created), so that queued a lookup for a
 * missing sheet on the shared request context. The switch then fell through,
 * the sheet was created correctly by the sheet handler, but the poisoned lookup
 * remained queued and the next ctx.sync() threw ItemNotFound — killing all
 * 100+ actions in the run.
 */
describe('handleWorksheetAction — does not touch worksheets it does not handle', () => {
  function makeCtx() {
    const getItem = vi.fn(() => ({}) as unknown);
    const getActiveWorksheet = vi.fn(() => ({}) as unknown);
    return {
      ctx: {
        workbook: { worksheets: { getItem, getActiveWorksheet } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext,
      getItem,
      getActiveWorksheet,
    };
  }

  it('never resolves a worksheet for ADD_SHEET (the sheet does not exist yet)', async () => {
    const { ctx, getItem, getActiveWorksheet } = makeCtx();
    const action = {
      type: 'ADD_SHEET',
      name: 'January',
      sheetName: 'January',
    } as unknown as RichAction;

    const handled = await handleWorksheetAction(action, ctx);

    expect(handled).toBe(false);
    // The critical assertion: no lookup was queued for the not-yet-created sheet.
    expect(getItem).not.toHaveBeenCalled();
    expect(getActiveWorksheet).not.toHaveBeenCalled();
  });

  it.each(['CREATE_SHEET', 'SET_CELL', 'BATCH_SET', 'SET_FORMULA', 'CREATE_CHART'])(
    'does not resolve a worksheet for unhandled type %s',
    async (type) => {
      const { ctx, getItem, getActiveWorksheet } = makeCtx();
      const action = { type, sheetName: 'February' } as unknown as RichAction;

      const handled = await handleWorksheetAction(action, ctx);

      expect(handled).toBe(false);
      expect(getItem).not.toHaveBeenCalled();
      expect(getActiveWorksheet).not.toHaveBeenCalled();
    },
  );

  it('still resolves and handles a type it owns', async () => {
    const { ctx, getItem } = makeCtx();
    const freezePanes = { freezeRows: vi.fn(), freezeColumns: vi.fn() };
    (ctx.workbook.worksheets.getItem as ReturnType<typeof vi.fn>).mockReturnValue({
      freezePanes,
    });

    const action = {
      type: 'FREEZE_PANES',
      sheetName: 'January',
      freezeRows: 1,
    } as unknown as RichAction;

    const handled = await handleWorksheetAction(action, ctx);

    expect(handled).toBe(true);
    expect(getItem).toHaveBeenCalledWith('January');
    expect(freezePanes.freezeRows).toHaveBeenCalledWith(1);
  });
});
