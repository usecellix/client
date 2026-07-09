import { ConversationHistoryMessage } from '@/utils/payloadCompressor';
import { ConversationTurn, truncateTabLabel } from '@/types/conversationTurn';
import { SheetAction } from '@/types/sheet-actions';

export interface StoredConversationMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  type?: 'question' | 'answer' | 'command' | 'clarification';
  timestamp?: string;
  metadata?: {
    actions?: unknown[];
    changeSetId?: string;
    questionOptions?: string[];
  };
}

export interface StoredConversation {
  conversationId: string;
  messages: StoredConversationMessage[];
  status?: string;
  updatedAt?: string;
}

export function messagesToHistory(messages: StoredConversationMessage[]): ConversationHistoryMessage[] {
  return messages.map((msg) => ({
    role: msg.role,
    content: msg.content,
    timestamp: msg.timestamp ?? new Date().toISOString(),
    type: msg.type ?? (msg.role === 'user' ? 'command' : 'answer'),
  }));
}

export function messagesToTurns(messages: StoredConversationMessage[]): ConversationTurn[] {
  const turns: ConversationTurn[] = [];

  for (const msg of messages) {
    if (msg.role === 'user') {
      const timestamp = msg.timestamp ? new Date(msg.timestamp) : new Date();
      turns.push({
        id: msg.id || `turn_${turns.length}_${Date.now()}`,
        userMessage: msg.content,
        timestamp,
        tabLabel: truncateTabLabel(msg.content),
        phase: 'complete',
        blocks: [],
      });
      continue;
    }

    const turn = turns[turns.length - 1];
    if (!turn) continue;

    if (msg.type === 'question' || msg.type === 'clarification') {
      turn.blocks.push({
        id: `question_${msg.id}`,
        type: 'question',
        question: msg.content.replace(/^\[Clarification needed\]: /, ''),
        options: msg.metadata?.questionOptions,
        revealState: 'visible',
      });
      turn.phase = 'awaiting_input';
      continue;
    }

    if (msg.metadata?.actions?.length) {
      turn.blocks.push({
        id: `actions_${msg.id}`,
        type: 'actions',
        actions: msg.metadata.actions as SheetAction[],
        explanation: msg.content,
        proposalStatus: 'accepted',
        changeSetId: msg.metadata.changeSetId,
      });
    }

    turn.blocks.push({
      id: `answer_${turn.id}`,
      type: 'answer',
      content: msg.content,
      revealState: 'complete',
    });
    turn.phase = 'complete';
  }

  return turns;
}

export function mergeSessionFromStored(
  localTurns: ConversationTurn[],
  localHistory: ConversationHistoryMessage[],
  stored: StoredConversation | null,
): { turns: ConversationTurn[]; history: ConversationHistoryMessage[] } {
  if (!stored?.messages?.length) {
    return { turns: localTurns, history: localHistory };
  }

  if (localTurns.length >= stored.messages.filter((m) => m.role === 'user').length) {
    return { turns: localTurns, history: localHistory };
  }

  return {
    turns: messagesToTurns(stored.messages),
    history: messagesToHistory(stored.messages),
  };
}
