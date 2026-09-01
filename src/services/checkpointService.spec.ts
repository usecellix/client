import { describe, it, expect, vi, afterEach } from 'vitest';
import { fetchCheckpoints, createCheckpoint, restoreCheckpoint } from './checkpointService';

/**
 * TASKS.md #31 — mirrors the fetch-mocking convention used for other frontend
 * service modules (stub global fetch, assert on the request shape and the
 * unwrapped response), same class of test workbookIdentity.spec.ts already
 * establishes for a different module in this repo.
 */

function mockFetchOnce(body: unknown, ok = true, status = 200) {
  const fn = vi.fn().mockResolvedValue({
    ok,
    status,
    statusText: ok ? 'OK' : 'Error',
    json: () => Promise.resolve(body),
    text: () => Promise.resolve(JSON.stringify(body)),
  });
  vi.stubGlobal('fetch', fn);
  return fn;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('checkpointService', () => {
  it('fetchCheckpoints GETs the workbook-scoped list endpoint and unwraps { checkpoints }', async () => {
    const fetchMock = mockFetchOnce({
      checkpoints: [
        {
          checkpointId: 'cp-1',
          workbookId: 'wb-1',
          conversationId: 'conv-1',
          label: 'auto',
          trigger: 'auto',
          anchorChangeSetId: 'cs-1',
          createdAt: '2026-08-19T00:00:00.000Z',
          status: 'active',
        },
      ],
    });

    const result = await fetchCheckpoints('wb-1');

    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining('/audit/checkpoint/wb-1'),
      expect.objectContaining({ credentials: 'include' }),
    );
    expect(result).toHaveLength(1);
    expect(result[0].checkpointId).toBe('cp-1');
  });

  it('fetchCheckpoints returns [] when the response has no checkpoints field', async () => {
    mockFetchOnce({});
    const result = await fetchCheckpoints('wb-1');
    expect(result).toEqual([]);
  });

  it('createCheckpoint POSTs workbookId/conversationId/label and returns the created checkpoint', async () => {
    const fetchMock = mockFetchOnce({
      checkpoint: {
        checkpointId: 'cp-2',
        workbookId: 'wb-1',
        conversationId: 'conv-1',
        label: 'before risky edit',
        trigger: 'manual',
        anchorChangeSetId: 'cs-1',
        createdAt: '2026-08-19T00:00:00.000Z',
        status: 'active',
      },
    });

    const checkpoint = await createCheckpoint('wb-1', 'conv-1', 'before risky edit');

    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining('/audit/checkpoint'),
      expect.objectContaining({
        method: 'POST',
        body: JSON.stringify({ workbookId: 'wb-1', conversationId: 'conv-1', label: 'before risky edit' }),
      }),
    );
    expect(checkpoint.checkpointId).toBe('cp-2');
    expect(checkpoint.trigger).toBe('manual');
  });

  it('restoreCheckpoint POSTs to the restore endpoint and returns checkpoint/revertedChangeSetIds/inverseActions', async () => {
    const fetchMock = mockFetchOnce({
      checkpoint: {
        checkpointId: 'cp-1',
        workbookId: 'wb-1',
        conversationId: 'conv-1',
        label: 'auto',
        trigger: 'auto',
        anchorChangeSetId: 'cs-1',
        createdAt: '2026-08-19T00:00:00.000Z',
        status: 'restored',
      },
      revertedChangeSetIds: ['cs-3', 'cs-2'],
      inverseActions: [{ type: 'SET_CELL', sheetName: 'Sales', address: 'A2', value: 'orig1' }],
    });

    const result = await restoreCheckpoint('cp-1');

    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining('/audit/checkpoint/restore/cp-1'),
      expect.objectContaining({ method: 'POST' }),
    );
    expect(result.checkpoint.status).toBe('restored');
    expect(result.revertedChangeSetIds).toEqual(['cs-3', 'cs-2']);
    expect(result.inverseActions).toHaveLength(1);
  });

  it('throws with the response status/body text on a non-OK response (e.g. #29 fail-closed 422)', async () => {
    mockFetchOnce(
      { message: 'Restore refused — cs-2 cannot be safely reverted', code: 'RESTORE_VERIFICATION_FAILED' },
      false,
      422,
    );
    await expect(restoreCheckpoint('cp-1')).rejects.toThrow('Checkpoint API 422');
  });
});
