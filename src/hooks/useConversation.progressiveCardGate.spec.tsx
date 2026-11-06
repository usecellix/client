// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useConversation } from '@/hooks/useConversation';
import { isTurnPresentationComplete } from '@/utils/turnPresentation';
import type { WorkbookContext } from '@/types/cellix.types';
import type { ActionBlock } from '@/types/conversationTurn';

/**
 * Live incident (5 Nov 2026): "Put the number 1 in cells A1 to A5" ran a
 * progressive Tier-3 build. The server finished in ~7s and sent
 * `actions` → … → `conversation_end` with NO `answer` and NO `wave_ready`.
 * Nothing opened the response gate, so the turn stayed in 'processing', the
 * Accept card stayed hidden and only pressing Stop revealed it.
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

const CTX = { activeSheet: 'Sheet2', sheets: ['Sheet2'] } as unknown as WorkbookContext;

const PROGRESSIVE_RUN = [
  sseBlock('status', { message: 'Working on your request…' }),
  sseBlock('thinking', { message: 'Planning your request...' }),
  sseBlock('status', { message: '1 step planned' }),
  sseBlock('status', { message: 'Prepared BATCH_SET on Sheet2' }),
  sseBlock('actions', {
    actions: [
      {
        type: 'BATCH_SET',
        sheetName: 'Sheet2',
        operations: ['A1', 'A2', 'A3', 'A4', 'A5'].map((address) => ({ address, value: 1 })),
      },
    ],
    explanation: 'Set the value 1 in cells A1 through A5',
    changeSetId: 'cs_progressive_1',
    changes: [{ cell: 'A1', sheet: 'Sheet2', before: '', after: 1 }],
    irreversibleActionTypes: [],
    tier: 3,
    stepLabel: 'Write content on Sheet2',
    progressive: true,
  }),
  sseBlock('status', { message: 'Progress: 1 step(s) ready · 1 change(s) prepared — continuing…' }),
  sseBlock('thinking', { message: 'Running deterministic checks...' }),
  sseBlock('status', { message: 'Actions verified' }),
  sseBlock('thinking', { message: 'Checks passed — preparing preview…' }),
  sseBlock('status', { message: '1 actions ready for preview' }),
  sseBlock('conversation_end', { summary: 'Review changes and accept or reject.', tier: 3 }),
  sseBlock('credits', { planCredits: 0, purchasedCredits: 0, oneTimeCredits: 112, debited: 2 }),
];

describe('useConversation — progressive run that ends with conversation_end only', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('finishes the turn and shows the Accept card without the user pressing Stop', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(makeSseResponse(PROGRESSIVE_RUN)));
    const onPreviewActions = vi.fn().mockResolvedValue(undefined);
    const onActions = vi.fn().mockResolvedValue(undefined);
    const { result } = renderHook(() =>
      useConversation({ workbookKey: 'progressive-gate', onActions, onClearPreview: vi.fn(), onPreviewActions }),
    );

    // sendMessage must resolve on its own — before the fix it never did.
    await act(async () => {
      await result.current.sendMessage('Put the number 1 in cells A1 to A5', [], CTX, 'Sheet2 is empty.', {
        mode: 'action',
      });
    });

    await waitFor(() => expect(result.current.turns).toHaveLength(1));
    await waitFor(() => expect(result.current.turns[0].phase).toBe('complete'));

    const turn = result.current.turns[0];

    // The answer block's 'typing' → 'complete' flip is done by the <AnswerReveal>
    // component on its first 24ms tick (not rendered in this hook-level test), so
    // apply it here. Everything else the hook controls — thinking, steps, pulsing
    // status — must already be finished, otherwise the card stays hidden.
    const afterTyping = {
      ...turn,
      blocks: turn.blocks.map((b) => (b.type === 'answer' ? { ...b, revealState: 'complete' as const } : b)),
    };
    expect(isTurnPresentationComplete(afterTyping)).toBe(true);

    const card = turn.blocks.find((b): b is ActionBlock => b.type === 'actions');
    expect(card?.proposalStatus).toBe('pending');
    // Preview only — nothing is applied until the user clicks Accept.
    expect(onPreviewActions).toHaveBeenCalled();
    expect(onActions).not.toHaveBeenCalled();
  }, 20000);
});
