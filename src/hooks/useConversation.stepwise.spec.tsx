// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useConversation } from '@/hooks/useConversation';
import { isTurnPresentationComplete } from '@/utils/turnPresentation';
import type { WorkbookContext } from '@/types/cellix.types';
import type { ActionBlock, ThinkingBlock } from '@/types/conversationTurn';

/**
 * Step-wise Tier 3 execution — TASKS.md #153, STEPWISE_EXECUTION.md.
 *
 * The load-bearing client guarantee: a card marked `stepwise` belongs to a run
 * that is PAUSED. Deciding it is the only thing that makes the next wave exist,
 * so an accept/reject that forgets to call `/continue` leaves the build stalled
 * forever with no error anywhere — exactly the silent-stall failure this suite
 * exists to prevent.
 */

function sseBlock(type: string, data: unknown): string {
  return `event: ${type}\ndata: ${JSON.stringify(data)}\n\n`;
}

function makeSseResponse(blocks: string[]): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const block of blocks) {
        controller.enqueue(encoder.encode(block));
      }
      controller.close();
    },
  });
  return { ok: true, status: 200, body: stream, headers: new Headers() } as unknown as Response;
}

const WORKBOOK_CONTEXT = { activeSheet: 'Sheet1', sheets: ['Sheet1'] } as unknown as WorkbookContext;
const PROMPT_CONTEXT = 'Sheet1 is empty.';
const SIMPLE_CREATE_MESSAGE = 'Create a table with sample headers and rows';

/** Wave 0's card plus the `wave_ready` that pauses the run. */
const WAVE_ZERO = [
  sseBlock('actions', {
    actions: [{ type: 'ADD_SHEET', sheetName: 'January' }],
    explanation: 'Create the month sheets',
    changeSetId: 'cs_wave_0',
    stepIndex: 1,
    stepTotal: 3,
    stepLabel: 'Create the month sheets',
    stepwise: true,
    runId: 'run_abc',
  }),
  sseBlock('wave_ready', {
    runId: 'run_abc',
    waveIndex: 0,
    waveTotal: 3,
    hasMore: true,
    changeSetId: 'cs_wave_0',
  }),
];

describe('useConversation — step-wise run continuation', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  async function sendStepwise(workbookKey: string, fetchMock: ReturnType<typeof vi.fn>) {
    vi.stubGlobal('fetch', fetchMock);
    const { result } = renderHook(() =>
      useConversation({
        workbookKey,
        // Accept applies via onActions; a no-op keeps the test on the hook's own
        // logic rather than Office.js, which does not exist in jsdom.
        onActions: vi.fn().mockResolvedValue(undefined),
      }),
    );
    await act(async () => {
      await result.current.sendMessage(SIMPLE_CREATE_MESSAGE, [], WORKBOOK_CONTEXT, PROMPT_CONTEXT, {
        mode: 'action',
      });
    });
    return result;
  }

  it('marks a paused wave card with its run id so Accept can advance the build', async () => {
    const fetchMock = vi.fn().mockResolvedValue(makeSseResponse(WAVE_ZERO));
    const result = await sendStepwise('stepwise-marks', fetchMock);

    await waitFor(() => expect(result.current.turns).toHaveLength(1));
    const block = result.current.turns[0].blocks.find(
      (b): b is ActionBlock => b.type === 'actions',
    );

    expect(block?.stepwise).toBe(true);
    expect(block?.runId).toBe('run_abc');
    expect(block?.proposalStatus).toBe('pending');
  });

  it('POSTs the accept decision to /continue, which is what generates the next wave', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
      .mockResolvedValueOnce(
        makeSseResponse([
          sseBlock('conversation_end', { summary: 'All steps applied.' }),
          sseBlock('wave_ready', {
            runId: 'run_abc',
            waveIndex: 1,
            waveTotal: 3,
            hasMore: false,
          }),
        ]),
      );

    const result = await sendStepwise('stepwise-accept', fetchMock);
    await waitFor(() => expect(result.current.turns).toHaveLength(1));

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

    await act(async () => {
      await result.current.acceptActions(turn.id, block.id);
    });

    // Second call is the continuation — without it the run stalls silently.
    expect(fetchMock).toHaveBeenCalledTimes(2);
    const [url, init] = fetchMock.mock.calls[1];
    expect(String(url)).toContain('/conversation/continue');
    const body = JSON.parse((init as RequestInit).body as string);
    expect(body).toMatchObject({ runId: 'run_abc', decision: 'accepted' });
  });

  it('continues the run on reject too, so a rejected step does not stall the build', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
      .mockResolvedValueOnce(
        makeSseResponse([
          sseBlock('conversation_end', { summary: 'Done — 1 step was not applied' }),
          sseBlock('wave_ready', {
            runId: 'run_abc',
            waveIndex: 1,
            waveTotal: 3,
            hasMore: false,
          }),
        ]),
      );

    const result = await sendStepwise('stepwise-reject', fetchMock);
    await waitFor(() => expect(result.current.turns).toHaveLength(1));

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

    await act(async () => {
      await result.current.rejectActions(turn.id, block.id);
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    const body = JSON.parse((fetchMock.mock.calls[1][1] as RequestInit).body as string);
    expect(body).toMatchObject({ runId: 'run_abc', decision: 'rejected' });
  });

  it('does not call /continue for a non-stepwise card', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      makeSseResponse([
        sseBlock('actions', {
          actions: [{ type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 5 }],
          explanation: 'Set A1',
          changeSetId: 'cs_plain',
        }),
        sseBlock('answer', { answer: 'Done.' }),
        sseBlock('conversation_end', {}),
      ]),
    );

    const result = await sendStepwise('stepwise-absent', fetchMock);
    await waitFor(() => expect(result.current.turns).toHaveLength(1));

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

    await act(async () => {
      await result.current.acceptActions(turn.id, block.id);
    });

    // One-shot runs have nothing to continue — a spurious /continue would 404.
    expect(fetchMock).toHaveBeenCalledTimes(1);
  }, 15000);

  it('a MIDDLE wave (arriving via /continue, hasMore: true) actually becomes presentable — not stuck invisible', async () => {
    // Live incident (Sept 9, 2026): a real build's second-of-two-step card was
    // generated correctly by the backend (confirmed in server logs) but never
    // appeared in the UI at all — no error, nothing in the DOM. Root cause:
    // `continueStepwiseRun` sets `phase: 'processing'` when a continuation
    // starts (so the wave's own live status updates can render), but the
    // `wave_ready` handler's `hasMore: true` branch used to only call
    // `signalResponse(...)` — which does nothing unless paired with
    // `revealFinalResponse`, reachable ONLY through `runVisualTimeline` (the
    // ORIGINAL send flow, never run by a continuation). So `phase` stayed
    // stuck at 'processing' forever, `isTurnPresentationComplete` returned
    // false, and `ConversationPanel` renders every PENDING action block as
    // `null` whenever the active turn isn't presentation-complete — the
    // card existed in state and never once reached the DOM.
    //
    // This is a THREE-wave scenario specifically because wave 0 goes through
    // `sendMessage`/`revealFinalResponse` (which already sets phase correctly
    // — this bug is invisible there) while wave 1 goes through
    // `continueStepwiseRun` alone, the exact path the bug lived in.
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
      .mockResolvedValueOnce(
        makeSseResponse([
          sseBlock('actions', {
            actions: [{ type: 'ADD_SHEET', sheetName: 'February' }],
            explanation: 'Create more month sheets',
            changeSetId: 'cs_wave_1',
            dependsOnChangeSetId: 'cs_wave_0',
            stepIndex: 2,
            stepTotal: 3,
            stepLabel: 'Create more month sheets',
            stepwise: true,
            runId: 'run_abc',
          }),
          sseBlock('wave_ready', {
            runId: 'run_abc',
            waveIndex: 1,
            waveTotal: 3,
            hasMore: true,
            changeSetId: 'cs_wave_1',
          }),
        ]),
      );

    const result = await sendStepwise('stepwise-middle-wave-visible', fetchMock);
    await waitFor(() => expect(result.current.turns).toHaveLength(1));

    const turnAfterWaveZero = result.current.turns[0];
    const waveZeroBlock = turnAfterWaveZero.blocks.find(
      (b): b is ActionBlock => b.type === 'actions',
    )!;

    await act(async () => {
      await result.current.acceptActions(turnAfterWaveZero.id, waveZeroBlock.id);
    });

    await waitFor(() => {
      const turn = result.current.turns[0];
      expect(turn.blocks.some((b) => b.type === 'actions' && b.id !== waveZeroBlock.id)).toBe(
        true,
      );
    });

    const turnAfterWaveOne = result.current.turns[0];
    // The actual end-to-end assertion: would ConversationPanel let this card
    // show its Accept button? Checking `phase` alone would have missed the
    // point of the bug — this is what genuinely gates visibility.
    expect(isTurnPresentationComplete(turnAfterWaveOne)).toBe(true);
    expect(turnAfterWaveOne.phase).toBe('complete');

    const waveOneBlock = turnAfterWaveOne.blocks.find(
      (b): b is ActionBlock => b.type === 'actions' && b.id !== waveZeroBlock.id,
    )!;
    expect(waveOneBlock.proposalStatus).toBe('pending');
  }, 15000);

  it('re-enters processing while a continuation is in flight, so progress updates actually render', async () => {
    // Live report (Sept 8, 2026): "no loading or anything" between steps.
    // Root cause — `TurnRenderer` hides every status/step/thinking block
    // while `turn.phase === 'complete'`, which it still was after the FIRST
    // wave's resolution. Real backend progress events during the SECOND
    // wave were written into turn state correctly but silently suppressed at
    // render time. This pins that the phase actually flips back — and then
    // forward again once the run genuinely finishes.
    //
    // A timer-based delay (not a manually-resolved Promise raced against a
    // concurrent waitFor) so the continuation's own async progression and
    // this test's polling both advance through the normal microtask/timer
    // queue, rather than the fragile "unawaited act() + concurrent waitFor"
    // pattern that made an earlier version of this test flaky.
    const fetchMock = vi.fn().mockResolvedValueOnce(makeSseResponse(WAVE_ZERO)).mockImplementationOnce(
      () =>
        new Promise<Response>((resolve) => {
          setTimeout(
            () =>
              resolve(
                makeSseResponse([
                  sseBlock('conversation_end', { summary: 'All steps applied.' }),
                  sseBlock('wave_ready', { runId: 'run_abc', waveIndex: 1, waveTotal: 3, hasMore: false }),
                ]),
              ),
            50,
          );
        }),
    );

    const result = await sendStepwise('stepwise-phase-reset', fetchMock);
    await waitFor(() => expect(result.current.turns).toHaveLength(1));
    expect(result.current.turns[0].phase).toBe('complete');

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

    void result.current.acceptActions(turn.id, block.id);

    await waitFor(() => {
      expect(result.current.turns[0].phase).toBe('processing');
    });

    // And the run genuinely finishing restores 'complete' — otherwise the
    // leftover "Preparing the next step…" status would sit on screen forever.
    await waitFor(() => {
      expect(result.current.turns[0].phase).toBe('complete');
    });
  }, 15000);

  it('keeps the thought process collapsed even when a continuation carries a build-completeness warning', async () => {
    // Live incident (Sept 9, 2026): the Planner truncated (again), and
    // TASKS.md #187's pruning correctly cut 7 of 10 subtasks rather than ship
    // broken formulas — then correctly told the user so via a 'status' event:
    // "I could not fully plan 7 steps of this request... ask again to add the
    // missing part". That text lands in the thinking log (never hidden once it
    // has content — TurnRenderer's own rule) but stays COLLAPSED unless
    // something marks it `expanded: true`. The one place that decided that —
    // `revealFinalResponse` — never runs for a continuation at all, so this
    // warning was technically present and functionally invisible: two
    // "Applied" cards, zero indication 70% of the request never got built.
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
      .mockResolvedValueOnce(
        makeSseResponse([
          sseBlock('status', {
            message:
              'Proceeding under an assumption — I could not fully plan 7 steps of this request (affecting Main, Lists) — the response was cut off before they were fully specified.',
          }),
          sseBlock('conversation_end', { summary: 'All steps applied.' }),
          sseBlock('wave_ready', { runId: 'run_abc', waveIndex: 1, waveTotal: 3, hasMore: false }),
        ]),
      );

    const result = await sendStepwise('stepwise-completeness-warning', fetchMock);
    await waitFor(() => expect(result.current.turns).toHaveLength(1));

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

    await act(async () => {
      await result.current.acceptActions(turn.id, block.id);
    });

    await waitFor(() => {
      const thinking = result.current.turns[0].blocks.find(
        (b): b is ThinkingBlock => b.type === 'thinking',
      );
      expect(thinking?.content).toMatch(/could not fully plan/i);
      // Present but collapsed — the thought process only opens on user tap.
      expect(thinking?.expanded).toBe(false);
    });
  }, 15000);

  it('leaves an ordinary thought process collapsed — no false positives from the widened warning check', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
      .mockResolvedValueOnce(
        makeSseResponse([
          sseBlock('status', { message: 'Populating the Lists sheet with dropdown values' }),
          sseBlock('conversation_end', { summary: 'All steps applied.' }),
          sseBlock('wave_ready', { runId: 'run_abc', waveIndex: 1, waveTotal: 3, hasMore: false }),
        ]),
      );

    const result = await sendStepwise('stepwise-no-false-positive', fetchMock);
    await waitFor(() => expect(result.current.turns).toHaveLength(1));

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

    await act(async () => {
      await result.current.acceptActions(turn.id, block.id);
    });

    await waitFor(() => {
      const thinking = result.current.turns[0].blocks.find(
        (b): b is ThinkingBlock => b.type === 'thinking',
      );
      expect(thinking?.content).toMatch(/Populating the Lists sheet/i);
      expect(thinking?.expanded).toBe(false);
    });
  }, 15000);

  it('keeps accepted work when a continuation fails, and says so', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
      .mockResolvedValueOnce({
        ok: false,
        status: 500,
        headers: new Headers({ 'content-type': 'text/plain' }),
        text: async () => 'run exploded',
        json: async () => ({}),
      } as unknown as Response);

    const result = await sendStepwise('stepwise-fail', fetchMock);
    await waitFor(() => expect(result.current.turns).toHaveLength(1));

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

    await act(async () => {
      await result.current.acceptActions(turn.id, block.id);
    });

    await waitFor(() => {
      const updated = result.current.turns[0];
      // The wave the user accepted stays accepted — a failed continuation must
      // never make already-applied work look un-applied.
      const acceptedBlock = updated.blocks.find((b): b is ActionBlock => b.type === 'actions');
      expect(acceptedBlock?.proposalStatus).toBe('accepted');
      expect(updated.error).toContain('still applied');
    });
  });
/**
   * TASKS.md #272 — a stepwise build spends essentially all its time in
   * /continue, and that fetch carried NO abort signal. Stop therefore could
   * not reach it, and because the request was never aborted the connection
   * stayed open, so the server-side cancellation keyed on the socket closing
   * (#260) could not fire either. Both halves were dead for long builds.
   */
  describe('stop + stall handling (TASKS.md #272)', () => {
    it('passes an abort signal on /continue, so Stop can actually reach the request', async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
        .mockResolvedValueOnce(
          makeSseResponse([
            sseBlock('conversation_end', { summary: 'All steps applied.' }),
            sseBlock('wave_ready', { runId: 'run_abc', waveIndex: 1, waveTotal: 3, hasMore: false }),
          ]),
        );

      const result = await sendStepwise('stepwise-abort-signal', fetchMock);
      await waitFor(() => expect(result.current.turns).toHaveLength(1));
      const turn = result.current.turns[0];
      const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

      await act(async () => {
        await result.current.acceptActions(turn.id, block.id);
      });

      const [, init] = fetchMock.mock.calls[1];
      const signal = (init as RequestInit).signal;
      expect(signal).toBeDefined();
      expect(signal).toBeInstanceOf(AbortSignal);
    });

    it('clears the waiting state after a wave, so the spinner cannot outlive the request', async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
        .mockResolvedValueOnce(
          makeSseResponse([
            sseBlock('wave_ready', { runId: 'run_abc', waveIndex: 1, waveTotal: 3, hasMore: true }),
          ]),
        );

      const result = await sendStepwise('stepwise-waiting-cleared', fetchMock);
      await waitFor(() => expect(result.current.turns).toHaveLength(1));
      const turn = result.current.turns[0];
      const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

      await act(async () => {
        await result.current.acceptActions(turn.id, block.id);
      });

      // A wave ends on wave_ready, which opens no response gate — the state
      // this asserts is exactly what left a spinning 'Thinking...' next to a
      // Send arrow, with no Stop button to press.
      await waitFor(() => expect(result.current.isWaitingForResponse).toBe(false));
    });

    it('a user Stop is not reported as a build failure', async () => {
      const abortError = new DOMException('The operation was aborted.', 'AbortError');
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(makeSseResponse(WAVE_ZERO))
        .mockRejectedValueOnce(abortError);

      const result = await sendStepwise('stepwise-user-stop', fetchMock);
      await waitFor(() => expect(result.current.turns).toHaveLength(1));
      const turn = result.current.turns[0];
      const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;

      await act(async () => {
        await result.current.acceptActions(turn.id, block.id);
      });

      await waitFor(() => {
        const updated = result.current.turns[0];
        // Stopping on purpose must not read as a crash.
        expect(updated.error).toBeUndefined();
        expect(result.current.isWaitingForResponse).toBe(false);
      });
    });
  });
});
