import { ChatSession, createChatSession } from '@/types/chatSession';
import { ConversationTurn, truncateTabLabel } from '@/types/conversationTurn';
import { ConversationHistoryMessage } from '@/utils/payloadCompressor';

export function appendUserTurnToSession(
  session: ChatSession,
  message: string,
  turnId: string,
  timestamp: Date = new Date(),
): ChatSession {
  const trimmed = message.trim();
  if (!trimmed) return session;

  const historyEntry: ConversationHistoryMessage = {
    role: 'user',
    content: trimmed,
    timestamp: timestamp.toISOString(),
    type: 'command',
  };

  const newTurn: ConversationTurn = {
    id: turnId,
    userMessage: trimmed,
    timestamp,
    tabLabel: truncateTabLabel(trimmed),
    phase: 'processing',
    blocks: [],
  };

  const nextTitle =
    session.turns.length === 0 ? truncateTabLabel(trimmed, 24) : session.title;

  return {
    ...session,
    title: nextTitle,
    updatedAt: timestamp.toISOString(),
    turns: [...session.turns, newTurn],
    history: [...session.history, historyEntry],
  };
}

export function addNewChatSession(sessions: ChatSession[]): {
  sessions: ChatSession[];
  activeSession: ChatSession;
} {
  const activeSession = createChatSession();
  return {
    sessions: [...sessions, activeSession],
    activeSession,
  };
}

export function setSessionConversationId(
  session: ChatSession,
  conversationId: string,
): ChatSession {
  return {
    ...session,
    conversationId,
    updatedAt: new Date().toISOString(),
  };
}
