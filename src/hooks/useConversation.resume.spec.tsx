// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useConversation } from '@/hooks/useConversation';

/**
 * Phase 8 resumability, client half — LONG_PROMPT_RELIABILITY_PLAN.md,
 * TASKS.md #293 (server) and #307 (this).
 *
 * Observed live: the task pane lost its connection mid-build, the run sat in
 * `awaiting_decision` forever, and nothing on either side could get back to
 * it. `findResumableRun` shipped and `getConversation` surfaced it — and
 * nothing consumed it, so the phase counted as done while the user still had
 * no way to continue a dropped build. These tests are what make it real.
 *
 * The load-bearing guarantee: resuming CONTINUES the existing run. Starting a
 * new one would rebuild sheets the user already accepted, which is worse than
 * the bug — that is why the assertions below check the endpoint and the run
 * id, not merely that something was sent.
 */

function sseBlock(type: string, data: unknown): string {
  return `event: ${type}\ndata: ${JSON.stringify(data)}\n\n`;
}

function makeSseResponse(blocks: string[]): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const block of blocks) controller.enqueue(encoder.encode(block));
      controller.close();
    },
  });
  return { ok: true, status: 200, body: stream, headers: new Headers() } as unknown as Response;
}

function jsonResponse(body: unknown): Response {
  return {
    ok: true,
    status: 200,
    json: async () => body,
    headers: new Headers(),
  } as unknown as Response;
}


/** A workbook whose stored session points at a conversation with a stranded run. */
function seedSession(workbookKey: string, conversationId: string): void {
  const now = new Date().toISOString();
  localStorage.setItem(
    `cellix:chat-sessions:${workbookKey}`,
    JSON.stringify({
      activeSessionId: 'session_seeded',
      sessions: [
        {
          id: 'session_seeded',
          conversationId,
          title: 'Twelve month build',
          createdAt: now,
          updatedAt: now,
          history: [],
          turns: [
            {
              id: 'turn_seeded',
              userMessage: 'Build twelve month sheets',
              timestamp: now,
              tabLabel: 'Build',
              phase: 'complete',
              blocks: [],
            },
          ],
        },
      ],
    }),
  );
}

describe('useConversation — resuming a dropped build (Phase 8)', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  const RESUMABLE = {
    conversationId: 'conv_stranded',
    messages: [],
    resumableRun: { runId: 'run_stranded', waveIndex: 3, waveTotal: 8 },
  };

  async function hydrate(workbookKey: string, fetchMock: ReturnType<typeof vi.fn>) {
    seedSession(workbookKey, 'conv_stranded');
    vi.stubGlobal('fetch', fetchMock);
    const { result } = renderHook(() =>
      useConversation({ workbookKey, onActions: vi.fn().mockResolvedValue(undefined) }),
    );
    return result;
  }

  it('surfaces the unfinished build when the conversation is reopened', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(RESUMABLE));
    const result = await hydrate('resume-surfaces', fetchMock);

    await waitFor(() => expect(result.current.resumableRun).not.toBeNull());
    expect(result.current.resumableRun).toEqual({
      runId: 'run_stranded',
      waveIndex: 3,
      waveTotal: 8,
    });
  });

  it('offers nothing when the conversation has no unfinished run', async () => {
    // The common case by far. An offer that appears on every reopen would be
    // noise, and noise gets dismissed reflexively — including when it matters.
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ conversationId: 'conv_stranded', messages: [] }));
    const result = await hydrate('resume-absent', fetchMock);

    await waitFor(() => expect(result.current.turns.length).toBeGreaterThanOrEqual(0));
    expect(result.current.resumableRun).toBeNull();
  });

  it('CONTINUES the existing run rather than starting a new build', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse(RESUMABLE))
      .mockResolvedValueOnce(
        makeSseResponse([
          sseBlock('conversation_end', { summary: 'Finished the remaining steps.' }),
          sseBlock('wave_ready', {
            runId: 'run_stranded',
            waveIndex: 7,
            waveTotal: 8,
            hasMore: false,
          }),
        ]),
      );

    const result = await hydrate('resume-continues', fetchMock);
    await waitFor(() => expect(result.current.resumableRun).not.toBeNull());

    await act(async () => {
      await result.current.resumeRun();
    });

    const continuation = fetchMock.mock.calls[1];
    expect(String(continuation[0])).toContain('/conversation/continue');
    const body = JSON.parse(String(continuation[1].body));
    expect(body.runId).toBe('run_stranded');
    expect(body.decision).toBe('accepted');
  });

  it('clears the offer once taken, so it cannot be continued twice', async () => {
    // A slow resume is exactly when someone clicks again. Two continuations of
    // one run would race each other through the same waves.
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse(RESUMABLE))
      .mockResolvedValue(
        makeSseResponse([sseBlock('conversation_end', { summary: 'Done.' })]),
      );

    const result = await hydrate('resume-once', fetchMock);
    await waitFor(() => expect(result.current.resumableRun).not.toBeNull());

    await act(async () => {
      await result.current.resumeRun();
    });

    expect(result.current.resumableRun).toBeNull();

    // A second call is a no-op: no further request goes out.
    const callsAfterFirst = fetchMock.mock.calls.length;
    await act(async () => {
      await result.current.resumeRun();
    });
    expect(fetchMock.mock.calls.length).toBe(callsAfterFirst);
  });

  it('dismissing leaves the run alone — it does not cancel anything', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(RESUMABLE));
    const result = await hydrate('resume-dismiss', fetchMock);
    await waitFor(() => expect(result.current.resumableRun).not.toBeNull());

    const callsBefore = fetchMock.mock.calls.length;
    act(() => {
      result.current.dismissResumableRun();
    });

    expect(result.current.resumableRun).toBeNull();
    expect(fetchMock.mock.calls.length).toBe(callsBefore);
  });

  it('a hydration failure never costs the conversation', async () => {
    // This runs on every reopen. It must degrade to "no offer", never throw.
    const fetchMock = vi.fn().mockRejectedValue(new Error('offline'));
    const result = await hydrate('resume-offline', fetchMock);

    await waitFor(() => expect(result.current.sessions.length).toBeGreaterThan(0));
    expect(result.current.resumableRun).toBeNull();
  });
});
