import { describe, expect, it } from 'vitest';
import { orderBlocksForDisplay } from './TurnRenderer';
import { ActionBlock, StatusBlock, ThinkingBlock, TurnBlock } from '@/types/conversationTurn';

/**
 * Live report: in a stepwise (TASKS.md #153) multi-wave turn, a NEW
 * "Thinking... Step 21" status/thinking block appended by a later wave
 * rendered ABOVE the already-Applied Step 1/Step 2 action cards instead of
 * below them — "the thinking part show it after the accept dialoage" (sic).
 * `blockPresentationOrder` unconditionally hoists every status/thinking
 * block above every actions block, which is correct for a single-shot turn
 * (progress always precedes its one resulting card) but wrong once a SECOND
 * wave's progress block arrives chronologically after an EARLIER wave's
 * already-rendered card — the reader is looking at "what happens next",
 * which belongs below what already happened, not above it.
 */

function status(id: string, label: string): StatusBlock {
  return { id, type: 'status', label, pulsing: true, visible: true };
}

function thinking(id: string, content: string): ThinkingBlock {
  return { id, type: 'thinking', content, expanded: false, visible: true };
}

function actionCard(id: string, stepIndex: number): ActionBlock {
  return {
    id,
    type: 'actions',
    actions: [],
    explanation: `Step ${stepIndex}`,
    proposalStatus: 'accepted',
    stepIndex,
    stepTotal: 6,
  };
}

describe('orderBlocksForDisplay', () => {
  it('keeps the FIRST wave\'s progress blocks hoisted above its own actions card (unchanged single-shot behavior)', () => {
    const blocks: TurnBlock[] = [
      status('s1', 'Planning your request...'),
      thinking('t1', 'Working out the details'),
      actionCard('a1', 1),
    ];

    const ordered = orderBlocksForDisplay(blocks);

    expect(ordered.map((b) => b.id)).toEqual(['s1', 't1', 'a1']);
  });

  it('renders a LATER wave\'s status/thinking block AFTER earlier waves\' action cards, not hoisted above them', () => {
    // Exact reported shape: two already-applied cards exist, then a new
    // "Thinking... Step 21" status block arrives for the next wave.
    const blocks: TurnBlock[] = [
      actionCard('a1', 1),
      actionCard('a2', 2),
      status('s2', 'Thinking... Step 21'),
      thinking('t2', 'Working out step 3'),
    ];

    const ordered = orderBlocksForDisplay(blocks);

    expect(ordered.map((b) => b.id)).toEqual(['a1', 'a2', 's2', 't2']);
  });

  it('still hoists a FIRST-wave status/thinking block above the SAME wave\'s own action card even when later blocks exist', () => {
    const blocks: TurnBlock[] = [
      status('s1', 'Planning your request...'),
      thinking('t1', 'Sketching the plan'),
      actionCard('a1', 1),
      status('s2', 'Thinking... Step 2'),
      actionCard('a2', 2),
    ];

    const ordered = orderBlocksForDisplay(blocks);

    // s1/t1 arrived BEFORE any actions block exists — still hoisted above a1.
    // s2 arrived AFTER a1 already existed — anchored after it, before a2
    // (its own true chronological position), not hoisted to the very top.
    expect(ordered.map((b) => b.id)).toEqual(['s1', 't1', 'a1', 's2', 'a2']);
  });

  it('does not reorder anything when there are no actions blocks at all', () => {
    const blocks: TurnBlock[] = [status('s1', 'Planning...'), thinking('t1', 'Thinking')];

    const ordered = orderBlocksForDisplay(blocks);

    expect(ordered.map((b) => b.id)).toEqual(['s1', 't1']);
  });
});
