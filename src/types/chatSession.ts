import { ConversationHistoryMessage } from '@/utils/payloadCompressor';
import { ConversationTurn } from '@/types/conversationTurn';

export interface ChatSession {
  id: string;
  title: string;
  conversationId: string | null;
  turns: ConversationTurn[];
  history: ConversationHistoryMessage[];
  createdAt: string;
  updatedAt: string;
}

export interface ChatSessionStore {
  activeSessionId: string | null;
  sessions: ChatSession[];
}

export function createChatSession(title = 'New chat'): ChatSession {
  const now = new Date().toISOString();
  return {
    id: `sess_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    title,
    conversationId: null,
    turns: [],
    history: [],
    createdAt: now,
    updatedAt: now,
  };
}
