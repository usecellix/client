import { ChatSession, ChatSessionStore } from '@/types/chatSession';
import { ConversationTurn, TurnBlock } from '@/types/conversationTurn';

const STORAGE_PREFIX = 'cellix:chat-sessions:';

function storageKey(workbookKey: string): string {
  return `${STORAGE_PREFIX}${workbookKey}`;
}

const IMPORT_EXPIRED_NOTE =
  'This import was not accepted before the pane was closed. Attach the file again to import it.';

/**
 * The copy of an attachment import that is saved to browser storage, with its
 * rows removed.
 *
 * An import card carries every transaction of the user's bank statement in its
 * write action. That is needed to apply the card, and for nothing after: once
 * accepted the rows are in the workbook. Saving them would leave a second copy
 * of the statement in browser storage indefinitely, and a long statement is
 * large enough to exhaust the storage quota and stop every chat from saving.
 */
export function stripImportedRows(block: TurnBlock): TurnBlock {
  if (block.type !== 'actions' || !block.importSheetName) return block;
  return {
    ...block,
    actions: block.actions.map((action) =>
      action.type === 'WRITE_TABLE' ? { ...action, rows: [] } : action,
    ),
  };
}

/**
 * The other half of `stripImportedRows`: an import still pending when it was
 * saved comes back without its rows, so it can no longer be applied. It is
 * closed with a note instead of being left as a card whose Accept would create
 * an empty sheet.
 */
export function expireRowlessImport(block: TurnBlock): TurnBlock {
  if (block.type !== 'actions' || !block.importSheetName) return block;
  if (block.proposalStatus !== 'pending') return block;
  return {
    ...block,
    proposalStatus: 'rejected',
    userFacingSummary: block.userFacingSummary
      ? { ...block.userFacingSummary, supportingDetail: IMPORT_EXPIRED_NOTE }
      : { headline: block.explanation, supportingDetail: IMPORT_EXPIRED_NOTE },
  };
}

function reviveTurnDates(turn: ConversationTurn): ConversationTurn {
  return {
    ...turn,
    timestamp: new Date(turn.timestamp),
    phase: turn.phase === 'processing' ? 'complete' : turn.phase,
    blocks: turn.blocks.map((block) => {
      if (block.type === 'actions') {
        return expireRowlessImport(block);
      }
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
      assistantMode: parsed.assistantMode,
      gstPurchasePortalPreference: parsed.gstPurchasePortalPreference,
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
      assistantMode: store.assistantMode,
      gstPurchasePortalPreference: store.gstPurchasePortalPreference,
      sessions: store.sessions.map((session) => ({
        ...session,
        turns: session.turns.map((turn) => ({
          ...turn,
          blocks: turn.blocks.map(stripImportedRows),
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
