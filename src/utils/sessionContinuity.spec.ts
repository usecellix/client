import { describe, expect, it } from 'vitest';
import { createChatSession } from '@/types/chatSession';
import {
  addNewChatSession,
  appendUserTurnToSession,
  setSessionConversationId,
} from '@/utils/sessionContinuity';

describe('sessionContinuity', () => {
  it('appends two turns to the same session and grows history', () => {
    const base = createChatSession();
    const afterFirst = appendUserTurnToSession(base, 'Find 2290', 'turn_1');
    const afterSecond = appendUserTurnToSession(afterFirst, 'Find 4180', 'turn_2');

    expect(afterSecond.id).toBe(base.id);
    expect(afterSecond.turns).toHaveLength(2);
    expect(afterSecond.turns[0]?.userMessage).toBe('Find 2290');
    expect(afterSecond.turns[1]?.userMessage).toBe('Find 4180');
    expect(afterSecond.history).toHaveLength(2);
    expect(afterSecond.history[0]?.content).toBe('Find 2290');
    expect(afterSecond.history[1]?.content).toBe('Find 4180');
    expect(afterSecond.title).toBe('Find 2290');
  });

  it('creates a new chat session with empty history and no conversation id', () => {
    const existing = appendUserTurnToSession(
      setSessionConversationId(createChatSession(), 'conv_existing'),
      'First prompt',
      'turn_1',
    );
    const { sessions, activeSession } = addNewChatSession([existing]);

    expect(sessions).toHaveLength(2);
    expect(activeSession.turns).toHaveLength(0);
    expect(activeSession.history).toHaveLength(0);
    expect(activeSession.conversationId).toBeNull();
    expect(sessions[0]?.conversationId).toBe('conv_existing');
  });

  it('assigns conversation id on first server reply', () => {
    const session = appendUserTurnToSession(createChatSession(), 'Hello', 'turn_1');
    const withId = setSessionConversationId(session, 'conv_new_1');

    expect(withId.conversationId).toBe('conv_new_1');
    expect(withId.history).toHaveLength(1);
    expect(withId.turns).toHaveLength(1);
  });
});
