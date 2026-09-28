import { ConversationHistoryMessage } from '@/utils/payloadCompressor';
import { ConversationTurn } from '@/types/conversationTurn';
import { AssistantMode } from '@/types/mode';

export interface ChatSession {
  id: string;
  title: string;
  conversationId: string | null;
  turns: ConversationTurn[];
  history: ConversationHistoryMessage[];
  createdAt: string;
  updatedAt: string;
}

/** Which GST portal source(s) a casual Purchase Register reconciliation should use when both GSTR-2B and GSTR-2A sheets are present in the workbook. */
export type GstPurchasePortalPreference = 'gstr2b_only' | 'gstr2a_only' | 'combined';

export interface ChatSessionStore {
  activeSessionId: string | null;
  sessions: ChatSession[];
  /** Last-selected assistant mode for this workbook (Ask / Plan / Action). */
  assistantMode?: AssistantMode;
  /**
   * Remembered per-workbook answer to "GSTR-2B only, GSTR-2A only, or combined?" —
   * asked once when both portal sheets are first seen together, so the user is never
   * asked again for this workbook unless they explicitly change it.
   */
  gstPurchasePortalPreference?: GstPurchasePortalPreference;
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
