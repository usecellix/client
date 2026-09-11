// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useConversation } from '@/hooks/useConversation';
import type { WorkbookContext } from '@/types/cellix.types';
import type { AssistantMode } from '@/types/mode';

/**
 * Interactive test infrastructure for `useConversation` — TASKS.md #63.
 *
 * The hook drives real state transitions over an async SSE stream (fetch +
 * ReadableStream + setTimeout-paced reveal animation), which needs a DOM and
 * `act()`-driven re-renders — `renderToStaticMarkup` (used elsewhere in this
 * repo) can't exercise it. The reveal pacing in `revealQueue.ts`'s TIMING
 * constants is short enough on the "simple create task" path (~450ms real
 * time) that real timers are used rather than fake ones — fake timers proved
 * unnecessary and, combined with React's scheduler, made the promise chain
 * hang rather than resolve (see session notes for the discarded attempt).
 */

function sseBlock(type: string, data: unknown): string {
  return `event: ${type}\ndata: ${JSON.stringify(data)}\n\n`;
}

/** A minimal `Response`-shaped mock whose `.body` streams the given SSE blocks. */
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
  return {
    ok: true,
    status: 200,
    body: stream,
    headers: new Headers(),
  } as unknown as Response;
}

const WORKBOOK_CONTEXT = { activeSheet: 'Sheet1', sheets: ['Sheet1'] } as unknown as WorkbookContext;
const PROMPT_CONTEXT = 'Sheet1 is empty.';
// isSimpleCreateTask() needs both an empty sheet and a create+row/column/table
// keyword match — keeps runVisualTimeline on its short "simple" branch so the
// test only waits out ~450ms of real reveal pacing, not the full multi-stage
// sequence (which totals several seconds per TIMING in revealQueue.ts).
const SIMPLE_CREATE_MESSAGE = 'Create a table with sample headers and rows';

async function sendAndSettle(
  result: { current: ReturnType<typeof useConversation> },
  mode: AssistantMode,
) {
  await act(async () => {
    await result.current.sendMessage(
      SIMPLE_CREATE_MESSAGE,
      [],
      WORKBOOK_CONTEXT,
      PROMPT_CONTEXT,
      { mode },
    );
  });
}

describe('useConversation — SSE-driven state transitions', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it(
    'never surfaces an Accept-able action block when the backend leaks actions during plan mode (TASKS.md #47)',
    async () => {
      vi.stubGlobal(
        'fetch',
        vi.fn().mockResolvedValue(
          makeSseResponse([
            sseBlock('actions', {
              actions: [{ type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 5 }],
              explanation: 'Set A1 to 5.',
              changeSetId: 'cs_leak_1',
              irreversibleActionTypes: [],
            }),
            sseBlock('answer', { answer: "I've planned this out — here's the approach." }),
            sseBlock('conversation_end', {}),
          ]),
        ),
      );

      const { result } = renderHook(() => useConversation({ workbookKey: 'plan-mode-test' }));

      await sendAndSettle(result, 'plan');

      await waitFor(() => expect(result.current.turns).toHaveLength(1));
      const turn = result.current.turns[0];

      const pendingActionBlock = turn.blocks.find(
        (b) => b.type === 'actions' && b.proposalStatus === 'pending',
      );
      expect(pendingActionBlock).toBeUndefined();

      // The read-only turn should still complete and surface the answer text —
      // this isn't just "nothing rendered," the actions specifically were dropped.
      expect(turn.phase).toBe('complete');
      expect(turn.blocks.some((b) => b.type === 'answer')).toBe(true);
    },
    10000,
  );

  it(
    'never surfaces an Accept-able action block during ask mode either',
    async () => {
      vi.stubGlobal(
        'fetch',
        vi.fn().mockResolvedValue(
          makeSseResponse([
            sseBlock('actions', {
              actions: [{ type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 5 }],
              explanation: 'Set A1 to 5.',
              changeSetId: 'cs_leak_2',
            }),
            sseBlock('answer', { answer: 'Sheet1!A1 is currently empty.' }),
            sseBlock('conversation_end', {}),
          ]),
        ),
      );

      const { result } = renderHook(() => useConversation({ workbookKey: 'ask-mode-test' }));

      await sendAndSettle(result, 'ask');

      await waitFor(() => expect(result.current.turns).toHaveLength(1));
      const turn = result.current.turns[0];

      expect(
        turn.blocks.find((b) => b.type === 'actions' && b.proposalStatus === 'pending'),
      ).toBeUndefined();
    },
    10000,
  );

  it(
    'sanity check: action mode DOES surface a pending action block for the same SSE payload',
    async () => {
      // Proves the harness itself is capable of producing a pending action
      // block — otherwise the two tests above could pass vacuously.
      vi.stubGlobal(
        'fetch',
        vi.fn().mockResolvedValue(
          makeSseResponse([
            sseBlock('actions', {
              actions: [{ type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 5 }],
              explanation: 'Set A1 to 5.',
              changeSetId: 'cs_real_1',
              irreversibleActionTypes: [],
            }),
            sseBlock('answer', { answer: "I've set A1 to 5." }),
            sseBlock('conversation_end', {}),
          ]),
        ),
      );

      const { result } = renderHook(() => useConversation({ workbookKey: 'action-mode-test' }));

      await sendAndSettle(result, 'action');

      await waitFor(() => expect(result.current.turns).toHaveLength(1));
      const turn = result.current.turns[0];

      const pendingActionBlock = turn.blocks.find(
        (b) => b.type === 'actions' && b.proposalStatus === 'pending',
      );
      expect(pendingActionBlock).toBeDefined();
    },
    10000,
  );

  it(
    'regenerateTurnId replaces the existing turn in place instead of appending a new one',
    async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(
          makeSseResponse([
            sseBlock('answer', { answer: 'First answer.' }),
            sseBlock('conversation_end', {}),
          ]),
        )
        .mockResolvedValueOnce(
          makeSseResponse([
            sseBlock('answer', { answer: 'Second answer, regenerated.' }),
            sseBlock('conversation_end', {}),
          ]),
        );
      vi.stubGlobal('fetch', fetchMock);

      const { result } = renderHook(() => useConversation({ workbookKey: 'regenerate-test' }));

      await sendAndSettle(result, 'ask');
      await waitFor(() => expect(result.current.turns).toHaveLength(1));
      const originalTurnId = result.current.turns[0].id;
      expect(
        result.current.turns[0].blocks.some(
          (b) => b.type === 'answer' && b.content === 'First answer.',
        ),
      ).toBe(true);

      await act(async () => {
        await result.current.sendMessage(
          SIMPLE_CREATE_MESSAGE,
          [],
          WORKBOOK_CONTEXT,
          PROMPT_CONTEXT,
          { mode: 'ask', regenerateTurnId: originalTurnId },
        );
      });

      // Same turn count, same id — not a second bubble appended below.
      expect(result.current.turns).toHaveLength(1);
      expect(result.current.turns[0].id).toBe(originalTurnId);
      await waitFor(() =>
        expect(
          result.current.turns[0].blocks.some(
            (b) => b.type === 'answer' && b.content === 'Second answer, regenerated.',
          ),
        ).toBe(true),
      );
    },
    10000,
  );

  it(
    'a greeting ("hi") skips the "Reading your worksheet…" choreography entirely — shows a single lightweight "Thinking…" step instead',
    async () => {
      // Mirrors the backend's real CHITCHAT route (conversation.service.ts's
      // handleChitchat): plain `chunk` streaming, no `status`/`thinking`/
      // `actions` events at all — reported live as "no need for that... show
      // thinking and show how do we handle it better."
      vi.stubGlobal(
        'fetch',
        vi.fn().mockResolvedValue(
          makeSseResponse([
            sseBlock('chunk', { text: 'Hi there! ' }),
            sseBlock('chunk', { text: 'What can I help you with?' }),
            sseBlock('conversation_end', {}),
          ]),
        ),
      );

      const { result } = renderHook(() => useConversation({ workbookKey: 'chitchat-test' }));

      const startedAt = Date.now();
      await act(async () => {
        await result.current.sendMessage('hi', [], WORKBOOK_CONTEXT, PROMPT_CONTEXT, {
          mode: 'ask',
        });
      });
      const elapsedMs = Date.now() - startedAt;

      await waitFor(() => expect(result.current.turns).toHaveLength(1));
      const turn = result.current.turns[0];

      // The actual regression: the full reading→analyzing pacing has hard
      // minimum delays (TIMING.readingMinRun + analyzingMinRun alone total
      // 4000ms, before the reveal/gap delays around them) that ran
      // regardless of how fast the real reply arrived. The chitchat fast
      // path's only wait is `waitWithMin(gate, 200)` — comfortably under 2s
      // even with test/CI slack, versus several multiples of that on the
      // full path (confirmed by disabling the fast path locally: the
      // identical scenario then took ~10s).
      expect(elapsedMs).toBeLessThan(2000);

      // Never shows the heavyweight worksheet-analysis labels — those imply
      // real workbook work that never happened for a bare greeting.
      expect(
        turn.blocks.some(
          (b) =>
            (b.type === 'step' || b.type === 'status') &&
            /reading your worksheet|analyzing your spreadsheet/i.test(b.label),
        ),
      ).toBe(false);
      expect(
        turn.blocks.some((b) => b.type === 'thinking' && /analyzing your spreadsheet/i.test(b.content)),
      ).toBe(false);

      expect(turn.phase).toBe('complete');
      expect(
        turn.blocks.some(
          (b) => b.type === 'answer' && b.content === 'Hi there! What can I help you with?',
        ),
      ).toBe(true);
    },
    10000,
  );
});
