import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  __resetPendingWorkbookContextForTests,
  getContextForSend,
  getPendingWorkbookSnapshot,
  markPendingWorkbookContextStale,
  setPendingWorkbookContext,
} from '@/utils/pendingWorkbookContext';

vi.mock('@/services/sheetContextBuilder', () => ({
  buildWorkbookContext: vi.fn(async () => ({
    context: { activeSheet: 'Sheet1', sheets: [] },
    activeSheetData: [['Header'], ['a']],
    promptContext: 'fresh-toon',
  })),
}));

import { buildWorkbookContext } from '@/services/sheetContextBuilder';

describe('pendingWorkbookContext (Spec 09 item 1)', () => {
  beforeEach(() => {
    __resetPendingWorkbookContextForTests();
    vi.mocked(buildWorkbookContext).mockClear();
  });

  it('reuses a fresh prebuild without calling buildWorkbookContext', async () => {
    setPendingWorkbookContext({
      toon: 'cached-toon',
      workbookContext: { activeSheet: 'Apps', sheets: [] },
      activeSheetData: [['H'], ['1']],
    });

    const first = await getContextForSend();
    const second = await getContextForSend();

    expect(first.reusedPending).toBe(true);
    expect(second.reusedPending).toBe(true);
    expect(first.promptContext).toBe('cached-toon');
    expect(second.promptContext).toBe('cached-toon');
    expect(buildWorkbookContext).not.toHaveBeenCalled();
  });

  it('rebuilds after markPendingWorkbookContextStale (external edit / apply)', async () => {
    setPendingWorkbookContext({
      toon: 'cached-toon',
      workbookContext: { activeSheet: 'Apps', sheets: [] },
      activeSheetData: [['H']],
    });
    markPendingWorkbookContextStale();
    expect(getPendingWorkbookSnapshot()?.stale).toBe(true);

    const resolved = await getContextForSend();
    expect(resolved.reusedPending).toBe(false);
    expect(resolved.promptContext).toBe('fresh-toon');
    expect(buildWorkbookContext).toHaveBeenCalledTimes(1);
    expect(getPendingWorkbookSnapshot()?.stale).toBe(false);
  });
});
