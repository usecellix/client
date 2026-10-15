import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ChatSession } from '@/types/chatSession';
import { ActionBlock, ConversationTurn, GstReconMissedRowsBlock, TurnBlock } from '@/types/conversationTurn';
import { loadChatSessions, loadChatSessionsClaimingLegacy, saveChatSessions, stripDecidedPayload } from '@/utils/chatSessionStorage';

/** A browser storage that refuses to hold more than `limit` characters in total, like a full localStorage. */
class LimitedStorage implements Storage {
  private store = new Map<string, string>();
  constructor(public limit: number) {}
  get length() {
    return this.store.size;
  }
  clear() {
    this.store.clear();
  }
  getItem(key: string) {
    return this.store.get(key) ?? null;
  }
  key(i: number) {
    return Array.from(this.store.keys())[i] ?? null;
  }
  removeItem(key: string) {
    this.store.delete(key);
  }
  setItem(key: string, value: string) {
    const others = [...this.store.entries()].filter(([k]) => k !== key).reduce((n, [, v]) => n + v.length, 0);
    if (others + value.length > this.limit) throw new DOMException('quota', 'QuotaExceededError');
    this.store.set(key, value);
  }
}

const KEY = 'demo.xlsx';
const bigRows = (n: number) => Array.from({ length: n }, (_, i) => [`2026-09-${(i % 28) + 1}`, 'x'.repeat(60), i, '', 'y'.repeat(40)]);

const reconCard = (id: string, status: ActionBlock['proposalStatus'], rows = 150): ActionBlock =>
  ({
    id,
    type: 'actions',
    proposalStatus: status,
    explanation: 'Create the bank reconciliation sheet?',
    gstReconOutputSheetName: 'Bank Recon Summary',
    actions: [
      { type: 'CREATE_SHEET', sheetName: 'Bank Recon Summary', name: 'Bank Recon Summary' },
      { type: 'WRITE_TABLE', sheetName: 'Bank Recon Summary', headers: ['A'], rows: bigRows(rows) },
    ],
  }) as unknown as ActionBlock;

const pointerList = (id: string, n = 60): GstReconMissedRowsBlock => ({
  id,
  type: 'gst_recon_missed_rows',
  rows: Array.from({ length: n }, (_, i) => ({
    vendorName: `Bank · ${i}`, gstin: '₹1 out', reason: 'r', explanation: 'e'.repeat(80), sheetName: null, row: null, inFile: true, pillLabel: `row ${i} in attached CSV`,
  })),
});

const turn = (id: string, blocks: TurnBlock[]): ConversationTurn =>
  ({ id, userMessage: `message ${id}`, timestamp: new Date('2026-10-14T10:00:00Z'), tabLabel: id, phase: 'complete', blocks }) as unknown as ConversationTurn;

const session = (id: string, turns: ConversationTurn[]): ChatSession => ({
  id, title: `chat ${id}`, conversationId: null, turns, history: [], createdAt: '2026-10-14T10:00:00Z', updatedAt: '2026-10-14T10:00:00Z',
});

const reconTurns = (prefix: string, n: number) =>
  Array.from({ length: n }, (_, i) => turn(`${prefix}${i}`, [reconCard(`${prefix}c${i}`, 'accepted'), pointerList(`${prefix}p${i}`)]));

describe('stripDecidedPayload', () => {
  it('drops the rows of a decided reconciliation card, keeps a pending one whole', () => {
    const decided = stripDecidedPayload(reconCard('a', 'accepted')) as ActionBlock;
    const rejected = stripDecidedPayload(reconCard('b', 'rejected')) as ActionBlock;
    const pending = stripDecidedPayload(reconCard('c', 'pending')) as ActionBlock;
    const rowsOf = (b: ActionBlock) => (b.actions.find((a) => a.type === 'WRITE_TABLE') as { rows: unknown[] }).rows.length;
    expect(rowsOf(decided)).toBe(0);
    expect(rowsOf(rejected)).toBe(0);
    expect(rowsOf(pending)).toBe(150); // still needed to apply it
  });

  it('leaves every other kind of card alone', () => {
    const ordinary = { id: 'x', type: 'actions', proposalStatus: 'accepted', explanation: 'e', actions: [{ type: 'WRITE_TABLE', sheetName: 'S', headers: [], rows: [[1]] }] } as unknown as ActionBlock;
    expect(stripDecidedPayload(ordinary)).toBe(ordinary);
  });
});

describe('saveChatSessions when the browser storage is limited', () => {
  let storage: LimitedStorage;
  beforeEach(() => {
    storage = new LimitedStorage(10_000_000);
    vi.stubGlobal('localStorage', storage);
    vi.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('keeps everything when it fits (nothing is trimmed needlessly)', () => {
    const sessions = [session('old', reconTurns('o', 3)), session('now', reconTurns('n', 3))];
    const result = saveChatSessions(KEY, { activeSessionId: 'now', sessions });
    expect(result).toEqual({ saved: true, tier: 0 });
    const loaded = loadChatSessions(KEY)!;
    expect(loaded.sessions.map((s) => s.id)).toEqual(['old', 'now']);
    expect(loaded.sessions[0].turns).toHaveLength(3);
  });

  it('a full storage can no longer lose the chat in use: older chats are trimmed to make room', () => {
    // Old chats hold 20 reconciliations each (decided cards are slimmed, but the pointer lists add up).
    const olds = Array.from({ length: 8 }, (_, k) => session(`old${k}`, reconTurns(`o${k}-`, 20)));
    const current = session('now', reconTurns('cur', 6));
    const sessions = [...olds, current];

    // Make the limit smaller than the full store, but easily big enough for the current chat alone.
    const full = JSON.stringify(sessions).length;
    storage.limit = Math.floor(full * 0.2);
    expect(() => storage.setItem('probe', 'x'.repeat(storage.limit + 1))).toThrow();

    const result = saveChatSessions(KEY, { activeSessionId: 'now', sessions });
    expect(result.saved).toBe(true);
    expect(result.tier).toBeGreaterThan(0);

    const loaded = loadChatSessions(KEY)!;
    expect(loaded.activeSessionId).toBe('now');
    const now = loaded.sessions.find((s) => s.id === 'now')!;
    expect(now.turns).toHaveLength(6); // the open chat is kept whole
    expect(now.title).toBe('chat now');
  });

  it('when only the open chat fits, that is what is saved (older chats go oldest-first)', () => {
    const olds = Array.from({ length: 5 }, (_, k) => session(`old${k}`, reconTurns(`o${k}-`, 15)));
    const current = session('now', reconTurns('cur', 10));

    // How much the open chat takes once saved (slimmed), measured with the real save.
    saveChatSessions(KEY, { activeSessionId: 'now', sessions: [current] });
    const currentOnly = storage.getItem(`cellix:chat-sessions:${KEY}`)!.length;
    storage.removeItem(`cellix:chat-sessions:${KEY}`);
    storage.limit = Math.floor(currentOnly * 1.3);

    const result = saveChatSessions(KEY, { activeSessionId: 'now', sessions: [...olds, current] });
    expect(result.saved).toBe(true);
    expect(result.tier).toBeGreaterThan(0);
    const loaded = loadChatSessions(KEY)!;
    const ids = loaded.sessions.map((s) => s.id);
    expect(ids).toContain('now');
    expect(loaded.sessions.find((s) => s.id === 'now')!.turns).toHaveLength(10); // whole
    expect(ids.length).toBeLessThan(6);
    // if any older chat survived, it is a recent one, not the oldest
    if (ids.length > 1) expect(ids).not.toContain('old0');
  });

  it('reports failure (instead of pretending) only when not even the open chat can be stored', () => {
    storage.limit = 50;
    const result = saveChatSessions(KEY, { activeSessionId: 'now', sessions: [session('now', reconTurns('cur', 3))] });
    expect(result.saved).toBe(false);
  });

  it('a pending card keeps its rows through a save (it still has to be applied)', () => {
    const sessions = [session('now', [turn('t', [reconCard('p', 'pending'), pointerList('l')])])];
    saveChatSessions(KEY, { activeSessionId: 'now', sessions });
    const card = loadChatSessions(KEY)!.sessions[0].turns[0].blocks.find((b) => b.type === 'actions') as ActionBlock;
    const write = card.actions.find((a) => a.type === 'WRITE_TABLE') as { rows: unknown[] };
    expect(write.rows).toHaveLength(150);
  });
});

/**
 * Browser storage is one budget for every workbook. The chat in use must still be saved when the space is
 * taken by other files' chats — the real case behind "my current chat is gone after a reload".
 */
describe('saveChatSessions when OTHER files\' chats fill the browser storage', () => {
  let storage: LimitedStorage;
  const OTHER = (n: string) => 'cellix:chat-sessions:' + n;
  const ownKey = 'cellix:chat-sessions:' + KEY;

  /** Another file's saved chats, written the old way: decided cards still carry all their sheet rows. */
  const writeOtherStore = (name: string, updatedAt: string, turns: number, rows = 300) => {
    const sessions = [
      { ...session('o-' + name, Array.from({ length: turns }, (_, i) => turn(name + i, [reconCard(name + 'c' + i, 'accepted', rows)]))), updatedAt },
    ].map((x) => ({ ...x, turns: x.turns.map((t) => ({ ...t, timestamp: t.timestamp.toISOString() })) }));
    storage.setItem(OTHER(name), JSON.stringify({ activeSessionId: 'o-' + name, sessions }));
  };

  beforeEach(() => {
    storage = new LimitedStorage(10_000_000);
    vi.stubGlobal('localStorage', storage);
    vi.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  const bytes = (k: string) => (storage.getItem(k) ?? '').length;

  it('step 1: other files\' chats are thinned out (nothing usable lost) and the current chat is saved', () => {
    writeOtherStore('a.xlsx', '2026-09-01T00:00:00Z', 6);
    writeOtherStore('b.xlsx', '2026-09-02T00:00:00Z', 6);
    const beforeA = bytes(OTHER('a.xlsx'));
    const current = session('now', reconTurns('cur', 4));

    // Room for the other files only once their rows are gone.
    const used = bytes(OTHER('a.xlsx')) + bytes(OTHER('b.xlsx'));
    storage.limit = used + 1000; // the storage is full: only thinning the other files makes room

    const result = saveChatSessions(KEY, { activeSessionId: 'now', sessions: [current] });
    expect(result.saved).toBe(true);
    expect(result.otherChats).toMatchObject({ slimmed: 2, removed: 0 });
    expect(loadChatSessions(KEY)!.sessions[0].turns).toHaveLength(4);
    // both other chats are still there, just lighter
    expect(storage.getItem(OTHER('a.xlsx'))).not.toBeNull();
    expect(bytes(OTHER('a.xlsx'))).toBeLessThan(beforeA);
    expect(loadChatSessions('a.xlsx')!.sessions[0].turns).toHaveLength(6);
  });

  it('step 2: when thinning is not enough, the least recently used other file\'s chats go first', () => {
    // Other chats made of plain text (nothing to thin out), oldest first.
    const heavyText = (name: string, updatedAt: string) => {
      const sessions = [{ ...session('o-' + name, [turn(name, [{ id: 'a', type: 'answer', content: 'x'.repeat(60_000), revealState: 'complete' } as TurnBlock])]), updatedAt }].map((x) => ({
        ...x,
        turns: x.turns.map((t) => ({ ...t, timestamp: t.timestamp.toISOString() })),
      }));
      storage.setItem(OTHER(name), JSON.stringify({ activeSessionId: 'o-' + name, sessions }));
    };
    heavyText('oldest.xlsx', '2026-08-01T00:00:00Z');
    heavyText('middle.xlsx', '2026-09-01T00:00:00Z');
    heavyText('newest.xlsx', '2026-10-01T00:00:00Z');
    const current = session('now', reconTurns('cur', 3));

    saveChatSessions('measure', { activeSessionId: 'now', sessions: [current] });
    const need = bytes('cellix:chat-sessions:measure');
    storage.removeItem('cellix:chat-sessions:measure');
    // Fits the current chat plus ONE of the three heavy stores, not two.
    storage.limit = need + bytes(OTHER('newest.xlsx')) + 5_000;

    const result = saveChatSessions(KEY, { activeSessionId: 'now', sessions: [current] });
    expect(result.saved).toBe(true);
    expect(result.otherChats?.removed).toBe(2);
    expect(storage.getItem(OTHER('oldest.xlsx'))).toBeNull();
    expect(storage.getItem(OTHER('middle.xlsx'))).toBeNull();
    expect(storage.getItem(OTHER('newest.xlsx'))).not.toBeNull(); // the most recently used survives
    expect(storage.getItem(ownKey)).not.toBeNull();
  });

  it('does not touch other files\' chats when the current chat fits on its own', () => {
    writeOtherStore('a.xlsx', '2026-09-01T00:00:00Z', 2);
    const before = storage.getItem(OTHER('a.xlsx'));
    const result = saveChatSessions(KEY, { activeSessionId: 'now', sessions: [session('now', reconTurns('cur', 2))] });
    expect(result).toEqual({ saved: true, tier: 0 });
    expect(storage.getItem(OTHER('a.xlsx'))).toBe(before);
  });
});

/**
 * Chats used to be saved under the placeholder name "workbook" while the real file name was unknown. After the
 * fix, a workbook loads from its own entry — the placeholder's history must not simply vanish.
 */
describe('loadChatSessionsClaimingLegacy', () => {
  let storage: LimitedStorage;
  beforeEach(() => {
    storage = new LimitedStorage(10_000_000);
    vi.stubGlobal('localStorage', storage);
    vi.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  const legacyStore = () =>
    saveChatSessions('workbook', {
      activeSessionId: 'bank',
      sessions: [session('gst', reconTurns('g', 2)), session('bank', reconTurns('b', 3))],
    });

  it('the first workbook opened adopts the placeholder\'s chats — moved, not copied', () => {
    legacyStore();
    // what your browser held: this workbook's own entry present but empty
    storage.setItem('cellix:chat-sessions:' + KEY, JSON.stringify({ activeSessionId: null, sessions: [] }));

    const claimed = loadChatSessionsClaimingLegacy(KEY)!;
    expect(claimed.sessions.map((x) => x.id)).toEqual(['gst', 'bank']);
    expect(claimed.activeSessionId).toBe('bank');
    expect(loadChatSessions(KEY)!.sessions).toHaveLength(2); // now saved under its own name
    expect(storage.getItem('cellix:chat-sessions:workbook')).toBeNull(); // and gone from the placeholder
  });

  it('a second workbook does not get the same chats again', () => {
    legacyStore();
    loadChatSessionsClaimingLegacy('first.xlsx');
    const second = loadChatSessionsClaimingLegacy('second.xlsx');
    expect(second?.sessions ?? []).toHaveLength(0);
  });

  it('a workbook that already has its own chats keeps them, and the placeholder is left alone', () => {
    legacyStore();
    saveChatSessions(KEY, { activeSessionId: 'mine', sessions: [session('mine', reconTurns('m', 1))] });
    const loaded = loadChatSessionsClaimingLegacy(KEY)!;
    expect(loaded.sessions.map((x) => x.id)).toEqual(['mine']);
    expect(storage.getItem('cellix:chat-sessions:workbook')).not.toBeNull();
  });

  it('does nothing when the real name could not be worked out (the placeholder IS the key)', () => {
    legacyStore();
    const loaded = loadChatSessionsClaimingLegacy('workbook')!;
    expect(loaded.sessions).toHaveLength(2);
    expect(storage.getItem('cellix:chat-sessions:workbook')).not.toBeNull();
  });

  it('never loses the placeholder\'s chats if they cannot be copied (storage full): they stay where they were', () => {
    legacyStore();
    storage.limit = storage.getItem('cellix:chat-sessions:workbook')!.length; // exactly full
    const loaded = loadChatSessionsClaimingLegacy(KEY);
    expect(loaded?.sessions.length ?? 0).toBeGreaterThan(0); // still shown
    expect(storage.getItem('cellix:chat-sessions:workbook')).not.toBeNull();
  });
});
