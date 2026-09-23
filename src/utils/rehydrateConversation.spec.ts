import { describe, expect, it } from 'vitest';
import { createChatSession } from '@/types/chatSession';
import {
  mergeSessionFromStored,
  messagesToHistory,
  messagesToTurns,
  StoredConversation,
} from '@/utils/rehydrateConversation';
import { appendUserTurnToSession } from '@/utils/sessionContinuity';

describe('rehydrateConversation', () => {
  it('maps stored messages to turns and history', () => {
    const stored: StoredConversation = {
      conversationId: 'conv_1',
      messages: [
        { id: 'u1', role: 'user', content: 'Find 2290', type: 'command' },
        { id: 'a1', role: 'assistant', content: 'Found 2290 in F30.', type: 'answer' },
        { id: 'u2', role: 'user', content: 'Find 4180', type: 'command' },
        { id: 'a2', role: 'assistant', content: 'Found 4180 in G12.', type: 'answer' },
      ],
    };

    const turns = messagesToTurns(stored.messages);
    const history = messagesToHistory(stored.messages);

    expect(turns).toHaveLength(2);
    expect(turns[0]?.blocks.some((block) => block.type === 'answer')).toBe(true);
    expect(turns[1]?.blocks.some((block) => block.type === 'answer')).toBe(true);
    expect(history).toHaveLength(4);
    expect(history[0]?.role).toBe('user');
    expect(history[1]?.role).toBe('assistant');
  });

  it('prefers local turns when they are at least as complete as stored messages', () => {
    let localSession = createChatSession();
    localSession = appendUserTurnToSession(localSession, 'Find 2290', 'turn_1');
    localSession = appendUserTurnToSession(localSession, 'Find 4180', 'turn_2');

    const stored: StoredConversation = {
      conversationId: 'conv_1',
      messages: [
        { id: 'u1', role: 'user', content: 'Find 2290', type: 'command' },
        { id: 'a1', role: 'assistant', content: 'Found 2290.', type: 'answer' },
      ],
    };

    const merged = mergeSessionFromStored(
      localSession.turns,
      localSession.history,
      stored,
    );

    expect(merged.turns).toHaveLength(2);
    expect(merged.history).toHaveLength(2);
  });

  it('fills from stored messages when local turns are behind', () => {
    const localSession = appendUserTurnToSession(createChatSession(), 'Find 2290', 'turn_1');
    const stored: StoredConversation = {
      conversationId: 'conv_1',
      messages: [
        { id: 'u1', role: 'user', content: 'Find 2290', type: 'command' },
        { id: 'a1', role: 'assistant', content: 'Found 2290.', type: 'answer' },
        { id: 'u2', role: 'user', content: 'Find 4180', type: 'command' },
        { id: 'a2', role: 'assistant', content: 'Found 4180.', type: 'answer' },
      ],
    };

    const merged = mergeSessionFromStored(
      localSession.turns,
      localSession.history,
      stored,
    );

    expect(merged.turns).toHaveLength(2);
    expect(merged.history).toHaveLength(4);
  });

  /**
   * TASKS.md #267 — a stepwise (multi-wave) build persists one assistant
   * message PER WAVE plus a closing summary, all inside the SAME turn (the
   * live case: a 12-sheet build across 7-9 waves). Every one of those
   * 'answer' blocks used to get the identical id `answer_${turn.id}`, which
   * collides as a React key in TurnRenderer's `key={block.id}` list — this
   * is what left a reopened multi-wave conversation showing an empty body
   * even after the server-persistence half of #267 was fixed.
   */
  it('gives each wave message its own answer block id within one multi-wave turn', () => {
    const stored: StoredConversation = {
      conversationId: 'conv_1',
      messages: [
        { id: 'u1', role: 'user', content: 'Build the ledger', type: 'command' },
        {
          id: 'wave1',
          role: 'assistant',
          content: 'Create sheets',
          type: 'answer',
          metadata: { actions: [{ type: 'ADD_SHEET' }], changeSetId: 'cs-1' },
        },
        {
          id: 'wave2',
          role: 'assistant',
          content: 'Write content on Main',
          type: 'answer',
          metadata: { actions: [{ type: 'BATCH_SET' }], changeSetId: 'cs-2' },
        },
        { id: 'wave3-final', role: 'assistant', content: 'All steps applied.', type: 'answer' },
      ],
    };

    const turns = messagesToTurns(stored.messages);

    expect(turns).toHaveLength(1);
    const answerBlocks = turns[0]!.blocks.filter((block) => block.type === 'answer');
    const actionBlocks = turns[0]!.blocks.filter((block) => block.type === 'actions');

    // The actual reported bug: every block must render (nothing lost to a
    // duplicate-key collision), and every id must be unique within the turn.
    expect(actionBlocks).toHaveLength(2);
    expect(answerBlocks).toHaveLength(3);
    const allIds = turns[0]!.blocks.map((block) => block.id);
    expect(new Set(allIds).size).toBe(allIds.length);
  });
});
