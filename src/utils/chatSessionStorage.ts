import { ChatSession, ChatSessionStore } from '@/types/chatSession';
import { ConversationTurn } from '@/types/conversationTurn';

const STORAGE_PREFIX = 'cellix:chat-sessions:';

function storageKey(workbookKey: string): string {
  return `${STORAGE_PREFIX}${workbookKey}`;
}

function reviveTurnDates(turn: ConversationTurn): ConversationTurn {
  return {
    ...turn,
    timestamp: new Date(turn.timestamp),
    phase: turn.phase === 'processing' ? 'complete' : turn.phase,
    blocks: turn.blocks.map((block) => {
      if (block.type === 'answer' && block.revealState === 'typing') {
        return { ...block, revealState: 'complete' as const };
      }
      if (block.type === 'thinking') {
        return { ...block, loading: false, visible: false, expanded: false };
      }
      if (block.type === 'status') {
        return { ...block, visible: false, pulsing: false };
      }
      return block;
    }),
  };
}

function reviveSession(session: ChatSession): ChatSession {
  return {
    ...session,
    turns: session.turns.map(reviveTurnDates),
  };
}

export function loadChatSessions(workbookKey: string): ChatSessionStore | null {
  if (!workbookKey) return null;
  try {
    const raw = localStorage.getItem(storageKey(workbookKey));
    if (!raw) return null;
    const parsed = JSON.parse(raw) as ChatSessionStore;
    if (!parsed || !Array.isArray(parsed.sessions)) return null;
    return {
      activeSessionId: parsed.activeSessionId ?? null,
      sessions: parsed.sessions.map(reviveSession),
    };
  } catch {
    return null;
  }
}

export function saveChatSessions(workbookKey: string, store: ChatSessionStore): void {
  if (!workbookKey) return;
  try {
    const payload: ChatSessionStore = {
      activeSessionId: store.activeSessionId,
      sessions: store.sessions.map((session) => ({
        ...session,
        turns: session.turns.map((turn) => ({
          ...turn,
          timestamp:
            turn.timestamp instanceof Date
              ? turn.timestamp.toISOString()
              : String(turn.timestamp),
        })) as unknown as ConversationTurn[],
      })),
    };
    localStorage.setItem(storageKey(workbookKey), JSON.stringify(payload));
  } catch (error) {
    console.warn('[Cellix] Failed to persist chat sessions:', error);
  }
}

export async function resolveWorkbookKey(): Promise<string> {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.load('name');
      await context.sync();
      const name = workbook.name?.trim();
      return name ? name.replace(/[^\w.-]+/g, '_') : 'workbook';
    });
  } catch {
    return 'workbook';
  }
}
