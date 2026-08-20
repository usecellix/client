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
});
