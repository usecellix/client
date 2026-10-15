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
 * A reconciliation or import card that has been accepted or rejected no longer needs the rows it would have
 * written: they are in the workbook (or were discarded). Only the pending ones keep them, because applying
 * the card needs them. One reconciliation summary is ~30 KB of write actions, so a long chat of runs would
 * otherwise use up the browser's storage and stop every chat from saving.
 */
export function stripDecidedPayload(block: TurnBlock): TurnBlock {
  if (block.type !== 'actions') return block;
  if (block.proposalStatus === 'pending') return block;
  if (!block.importSheetName && !block.gstReconOutputSheetName) return block;
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

/** How much of an OLDER chat is kept when the browser's storage is full. The chat in use is always kept whole. */
const OLDER_CHAT_TURNS_TIER1 = 12;
const OLDER_CHAT_TURNS_TIER2 = 5;
const ACTIVE_CHAT_TURNS_LAST_RESORT = 25;

function serializeSession(session: ChatSession): ChatSession {
  return {
    ...session,
    turns: session.turns.map((turn) => ({
      ...turn,
      blocks: turn.blocks.map((block) => stripDecidedPayload(stripImportedRows(block))),
      timestamp:
        turn.timestamp instanceof Date ? turn.timestamp.toISOString() : String(turn.timestamp),
    })) as unknown as ConversationTurn[],
  };
}

/** Drops everything bulky from the decided cards and the pointer lists of a chat (keeps the conversation itself). */
function compactSession(session: ChatSession, keepTurns: number): ChatSession {
  return {
    ...session,
    turns: session.turns.slice(-keepTurns).map((turn) => ({
      ...turn,
      blocks: turn.blocks.map((block) => {
        if (block.type === 'actions' && block.proposalStatus !== 'pending') {
          return { ...block, actions: [] };
        }
        if (block.type === 'gst_recon_missed_rows') return { ...block, rows: block.rows.slice(0, 10) };
        return block;
      }),
    })),
  };
}

/**
 * The sessions to write, from "everything" down to "just the chat in use". Each step is tried in turn until
 * the browser accepts one, so the chat the user is in is never the thing that gets lost: older chats are
 * trimmed first, then dropped oldest-first, and only as a last resort is the open chat itself shortened.
 */
function* storageTiers(sessions: ChatSession[], activeId: string | null): Generator<ChatSession[]> {
  const isActive = (session: ChatSession) => session.id === activeId;
  const older = sessions.filter((s) => !isActive(s));
  const active = sessions.filter(isActive);

  yield sessions;
  yield sessions.map((s) => (isActive(s) ? s : compactSession(s, OLDER_CHAT_TURNS_TIER1)));
  yield sessions.map((s) => (isActive(s) ? s : compactSession(s, OLDER_CHAT_TURNS_TIER2)));
  // Older chats go, oldest first, one at a time.
  for (let keep = older.length - 1; keep >= 0; keep--) {
    const kept = older.slice(older.length - keep).map((s) => compactSession(s, OLDER_CHAT_TURNS_TIER2));
    yield [...kept, ...active];
  }
  yield active.map((s) => compactSession(s, ACTIVE_CHAT_TURNS_LAST_RESORT));
}

export interface SaveResult {
  saved: boolean;
  /** 0 = everything was kept; higher = more of the older history had to be trimmed to fit. */
  tier: number;
  /** Other workbooks' saved chats that had to be thinned out (slimmed) or removed to make room. */
  otherChats?: { slimmed: number; removed: number };
}

/**
 * Browser storage is one budget for EVERY workbook's chats (and anything else the add-in keeps), so a full
 * budget can be caused by chats of other files. Trying to fit this chat by trimming only its own workbook's
 * history is not enough then.
 */
function otherChatKeys(exceptKey: string): string[] {
  const keys: string[] = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k && k.startsWith(STORAGE_PREFIX) && k !== exceptKey) keys.push(k);
  }
  return keys;
}

function lastTouched(raw: string | null): number {
  if (!raw) return 0;
  try {
    const parsed = JSON.parse(raw) as { sessions?: Array<{ updatedAt?: string }> };
    return Math.max(0, ...(parsed.sessions ?? []).map((s) => Date.parse(s.updatedAt ?? '') || 0));
  } catch {
    return 0;
  }
}

/** Rewrites another workbook's saved chats without the rows of already-decided cards. Nothing is lost that was still usable. */
function slimOtherChats(storeKey: string): boolean {
  try {
    const raw = localStorage.getItem(storeKey);
    if (!raw) return false;
    const parsed = JSON.parse(raw) as ChatSessionStore;
    if (!parsed || !Array.isArray(parsed.sessions)) return false;
    const slimmed = {
      ...parsed,
      sessions: parsed.sessions.map((session) => ({
        ...session,
        turns: session.turns.map((turn) => ({
          ...turn,
          blocks: (turn.blocks ?? []).map((block) => stripDecidedPayload(stripImportedRows(block))),
        })),
      })),
    };
    const next = JSON.stringify(slimmed);
    if (next.length >= raw.length) return false;
    localStorage.setItem(storeKey, next);
    return true;
  } catch {
    return false;
  }
}

/** The name chats were saved under before the real file name was known (the old placeholder). */
export const LEGACY_WORKBOOK_KEY = 'workbook';

/**
 * Loads a workbook's saved chats. Chats used to be loaded and saved under a placeholder name while the real
 * file name was still being worked out, so a workbook's own entry could be empty while its history sat under
 * the placeholder. The first workbook opened after that was fixed ADOPTS the placeholder's chats (moved, not
 * copied, so they appear once and are never lost); every later workbook starts from its own entry.
 */
export function loadChatSessionsClaimingLegacy(workbookKey: string): ChatSessionStore | null {
  const own = loadChatSessions(workbookKey);
  if (own?.sessions.length || !workbookKey || workbookKey === LEGACY_WORKBOOK_KEY) return own;

  const legacy = loadChatSessions(LEGACY_WORKBOOK_KEY);
  if (!legacy?.sessions.length) return own;

  const adopted: ChatSessionStore = {
    ...legacy,
    assistantMode: own?.assistantMode ?? legacy.assistantMode,
  };
  // Without making room in other chats: the placeholder entry is itself one of them, and is what is being moved.
  const { saved } = saveChatSessions(workbookKey, adopted, { makeRoomInOtherChats: false });
  if (!saved) return own ?? adopted; // show them now; the old entry stays so nothing is lost
  try {
    localStorage.removeItem(storageKey(LEGACY_WORKBOOK_KEY));
  } catch {
    /* harmless: it will simply not be offered again once this workbook has its own copy */
  }
  return loadChatSessions(workbookKey);
}

export interface SaveOptions {
  /**
   * When this chat does not fit, make room by thinning and then removing OTHER workbooks' saved chats.
   * Default true. Turned off by the one-time move of the old placeholder's chats, which must never delete
   * what it is in the middle of moving.
   */
  makeRoomInOtherChats?: boolean;
}

export function saveChatSessions(workbookKey: string, store: ChatSessionStore, options: SaveOptions = {}): SaveResult {
  if (!workbookKey) return { saved: false, tier: 0 };
  const makeRoom = options.makeRoomInOtherChats !== false;
  const ownKey = storageKey(workbookKey);
  const header = {
    activeSessionId: store.activeSessionId,
    assistantMode: store.assistantMode,
    gstPurchasePortalPreference: store.gstPurchasePortalPreference,
  };
  const prepared = store.sessions.map(serializeSession);

  let lastError: unknown;
  /** Tries every tier, from "everything" down to "just the open chat", until the browser accepts one. */
  const attempt = (): { saved: boolean; tier: number } => {
    let tier = 0;
    for (const sessions of storageTiers(prepared, store.activeSessionId)) {
      try {
        localStorage.setItem(ownKey, JSON.stringify({ ...header, sessions }));
        return { saved: true, tier };
      } catch (error) {
        lastError = error; // most often QuotaExceededError: try the next, smaller tier
        tier += 1;
      }
    }
    return { saved: false, tier };
  };

  let result = attempt();
  const otherChats = { slimmed: 0, removed: 0 };

  if (!result.saved && makeRoom) {
    // Step 1 — thin out the other workbooks' chats (non-destructive: only rows that were already applied or discarded).
    for (const k of otherChatKeys(ownKey)) if (slimOtherChats(k)) otherChats.slimmed += 1;
    if (otherChats.slimmed) result = attempt();
  }
  if (!result.saved && makeRoom) {
    // Step 2 — last resort: remove other workbooks' saved chats, least recently used first, until this chat fits.
    const victims = otherChatKeys(ownKey)
      .map((k) => ({ k, at: lastTouched(localStorage.getItem(k)) }))
      .sort((x, y) => x.at - y.at);
    for (const { k } of victims) {
      localStorage.removeItem(k);
      otherChats.removed += 1;
      result = attempt();
      if (result.saved) break;
    }
  }

  if (result.saved) {
    if (result.tier > 0 || otherChats.slimmed || otherChats.removed) {
      console.warn(
        `[Cellix] Browser storage is nearly full — made room so this chat could be saved (older history trimmed: level ${result.tier}; other files' chats thinned: ${otherChats.slimmed}, removed: ${otherChats.removed}).`,
      );
    }
    return { saved: true, tier: result.tier, ...(otherChats.slimmed || otherChats.removed ? { otherChats } : {}) };
  }
  console.warn('[Cellix] Failed to persist chat sessions:', lastError);
  return { saved: false, tier: result.tier, otherChats };
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
