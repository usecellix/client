import { describe, expect, it, vi } from 'vitest';
import { handleAddSheet } from './sheet.handler';
import { AddSheetAction } from '@/action.types';

/* global Excel */

/**
 * `created` flips true the moment `.add()` is called — modeling the real-world
 * failure mode this file guards against: Office.js's single batched sync() can
 * genuinely create the sheet on the document and still have the call reject
 * over an unrelated trailing op (activate/read-name), so "the sheet exists" and
 * "the wrapping sync() succeeded" are two different facts.
 */
function makeCtx(opts: { cosmeticSyncThrows?: boolean } = {}) {
  let created = false;
  let syncCallCount = 0;
  const addedNames: string[] = [];

  const worksheets = {
    getItemOrNullObject: vi.fn((name: string) => ({
      get isNullObject() {
        return !created;
      },
      name,
      load: vi.fn(),
      activate: vi.fn(),
    })),
    add: vi.fn((name: string) => {
      created = true;
      addedNames.push(name);
      return { name, load: vi.fn(), activate: vi.fn() };
    }),
  };

  const ctx = {
    workbook: { worksheets },
    sync: vi.fn(async () => {
      syncCallCount += 1;
      // Sync #1 = the initial existence check; sync #2 = the post-add cosmetic
      // sync (activate + load('name')) — the one this fix protects.
      if (opts.cosmeticSyncThrows && syncCallCount === 2) {
        throw new Error('OfficeExtension.Error: Error code 0x80070057');
      }
    }),
  };

  return { ctx: ctx as unknown as Excel.RequestContext, worksheets, addedNames, isCreated: () => created };
}

function addAction(name: string): AddSheetAction {
  return { type: 'ADD_SHEET', name };
}

describe('handleAddSheet', () => {
  it('creates a new sheet and returns its actual name (baseline, no error)', async () => {
    const { ctx } = makeCtx();
    const result = await handleAddSheet(addAction('Missed vs GSTR-2B'), ctx);
    expect(result).toEqual({
      requestedName: 'Missed vs GSTR-2B',
      actualName: 'Missed vs GSTR-2B',
      reusedExisting: false,
    });
  });

  it('reuses an existing sheet by activating it, never throwing (baseline)', async () => {
    const { ctx, worksheets } = makeCtx();
    // Pre-seed as already existing.
    (worksheets.getItemOrNullObject as ReturnType<typeof vi.fn>).mockImplementation((name: string) => ({
      isNullObject: false,
      name,
      load: vi.fn(),
      activate: vi.fn(),
    }));

    const result = await handleAddSheet(addAction('Missed vs GSTR-2B'), ctx);
    expect(result.reusedExisting).toBe(true);
  });

  it('a cosmetic post-creation sync failure does NOT surface as an ADD_SHEET error when the sheet was actually created — verified with a fresh existence check', async () => {
    const { ctx, isCreated } = makeCtx({ cosmeticSyncThrows: true });

    const result = await handleAddSheet(addAction('Missed vs GSTR-2B'), ctx);

    expect(isCreated()).toBe(true);
    expect(result.reusedExisting).toBe(false);
    expect(result.actualName).toBe('Missed vs GSTR-2B');
  });

  it('re-throws when the sync failure meant the sheet genuinely was not created', async () => {
    let syncCallCount = 0;
    const worksheets = {
      getItemOrNullObject: vi.fn((name: string) => ({
        isNullObject: true, // never created, in either the pre-check or the post-failure verify
        name,
        load: vi.fn(),
        activate: vi.fn(),
      })),
      add: vi.fn((name: string) => ({ name, load: vi.fn(), activate: vi.fn() })),
    };
    const ctx = {
      workbook: { worksheets },
      sync: vi.fn(async () => {
        syncCallCount += 1;
        if (syncCallCount === 2) {
          throw new Error('OfficeExtension.Error: Error code 0x80070057');
        }
      }),
    } as unknown as Excel.RequestContext;

    await expect(handleAddSheet(addAction('Missed vs GSTR-2B'), ctx)).rejects.toThrow(
      /0x80070057/,
    );
  });
});
