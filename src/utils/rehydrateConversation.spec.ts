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
});
