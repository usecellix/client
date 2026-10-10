import { afterEach, describe, expect, it, vi } from 'vitest';
import { markChangeSetApplied } from './auditService';
import type { FormatSnapshot } from '@/engine/formatSnapshot';

/** TASKS.md #400 — the apply call carries what the formatting overwrote, or Revert has nothing to work from. */
function stubFetch() {
  const fetchMock = vi.fn(async (_url: string, _init?: RequestInit) => ({
    ok: true,
    json: async () => ({ data: { changeSet: { changeSetId: 'cs-1' } } }),
  }));
  vi.stubGlobal('fetch', fetchMock);
  return fetchMock;
}

const snapshot: FormatSnapshot = {
  kind: 'format',
  sheetName: 'Monthly Summary',
  restorable: true,
  row: 2,
  col: 0,
  rowCount: 1,
  colCount: 4,
  palette: [{ bold: false, clearFill: true }],
  grid: [[0, 0, 0, 0]],
};

afterEach(() => vi.unstubAllGlobals());

describe('markChangeSetApplied — formatSnapshots', () => {
  it('sends the snapshots in the body', async () => {
    const fetchMock = stubFetch();
    await markChangeSetApplied('cs-1', undefined, undefined, undefined, [snapshot]);

    const init = fetchMock.mock.calls[0][1] as RequestInit;
    expect(init.method).toBe('POST');
    expect(JSON.parse(String(init.body))).toEqual({ formatSnapshots: [snapshot] });
  });

  it('sends an empty object when there is nothing to report, as before', async () => {
    const fetchMock = stubFetch();
    await markChangeSetApplied('cs-1', undefined, undefined, undefined, []);
    expect(String((fetchMock.mock.calls[0][1] as RequestInit).body)).toBe('{}');
  });
});
