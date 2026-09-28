import { describe, expect, it } from 'vitest';
import { orderBlocksForDisplay } from './TurnRenderer';
import {
  ActionBlock,
  GstReconMissedRowsBlock,
  QuestionBlock,
  StatusBlock,
  ThinkingBlock,
  TurnBlock,
} from '@/types/conversationTurn';

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

function question(id: string, answeredWith?: string): QuestionBlock {
  return {
    id,
    type: 'question',
    question: 'Where should the new row go?',
    options: ['After the last row'],
    ...(answeredWith ? { answeredWith } : {}),
  };
}

function missedRows(id: string): GstReconMissedRowsBlock {
  return { id, type: 'gst_recon_missed_rows', rows: [] };
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

  // TASKS.md #195 — answering continues the turn IN PLACE (#194), so the work
  // the answer kicks off arrives after the question block. `question` ranks 3
  // while status/thinking rank 1/2, which floated all of that progress back
  // above the card the user had just answered.
  it('keeps progress that follows an ANSWERED question below it', () => {
    const blocks: TurnBlock[] = [
      thinking('t1', 'First pass'),
      question('q1', 'After the last row (recommended)'),
      status('s2', 'Planning your request...'),
      thinking('t2', 'Second pass'),
    ];

    const ordered = orderBlocksForDisplay(blocks);

    expect(ordered.map((b) => b.id)).toEqual(['t1', 'q1', 's2', 't2']);
  });

  // An unanswered question anchors nothing: no work has followed it yet, and it
  // still belongs below the progress that produced it.
  it('leaves an UNANSWERED question below the progress that produced it', () => {
    const blocks: TurnBlock[] = [
      status('s1', 'Reading your worksheet...'),
      question('q1'),
    ];

    const ordered = orderBlocksForDisplay(blocks);

    expect(ordered.map((b) => b.id)).toEqual(['s1', 'q1']);
  });

  // Response actions (Accept/Reject) always read as the last thing in a turn — the GST
  // recon missed-rows pill list is reference material for the CA to consult before
  // deciding, so it must render just above the actions card, even though both are
  // appended to `turn.blocks` with the actions card pushed first (useConversation.ts's
  // applyGstReconOutcome builds the actions block, then the missed-rows block).
  it('renders the GST-recon missed-rows list ABOVE the actions card, even though the actions block was appended first', () => {
    const blocks: TurnBlock[] = [
      actionCard('a1', 1),
      missedRows('m1'),
    ];

    const ordered = orderBlocksForDisplay(blocks);

    expect(ordered.map((b) => b.id)).toEqual(['m1', 'a1']);
  });

  // Real report: after answering the "GSTR-2B or GSTR-2A?" question, the recon result's
  // "Compared 22 rows in Purchase Reg against GSTR-2B." answer text rendered AFTER the
  // missed-rows pill list instead of before it. Root cause: the answered-question anchor
  // made this `answer` block count as "progress continuing after an anchor" (TASKS.md
  // #195's fix), which bumped it to the SAME tier as the actions card — one tier BELOW
  // the missed-rows list's own tier — even though no `actions` block existed yet at the
  // point this `answer` block arrived. The fix distinguishes: a continuation is only
  // bumped to the actions tier when a REAL actions block already precedes it; here the
  // only earlier anchor is the answered question, so it must land just before the
  // missed-rows/actions tier instead, not tied with or after it.
  it('keeps the recon-result answer text ABOVE the missed-rows list and the actions card, when the only earlier anchor is an answered question (not yet an actions block)', () => {
    const blocks: TurnBlock[] = [
      question('q1', 'GSTR-2B only'),
      { id: 'ans1', type: 'answer', content: 'Compared 22 rows in Purchase Reg against GSTR-2B.', revealState: 'complete' },
      actionCard('a1', 1),
      missedRows('m1'),
    ];

    const ordered = orderBlocksForDisplay(blocks);

    expect(ordered.map((b) => b.id)).toEqual(['q1', 'ans1', 'm1', 'a1']);
  });
});
