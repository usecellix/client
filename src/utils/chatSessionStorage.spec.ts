import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createChatSession } from '@/types/chatSession';
import { appendUserTurnToSession } from '@/utils/sessionContinuity';
import { loadChatSessions, saveChatSessions } from '@/utils/chatSessionStorage';
import { ActionBlock, ConversationTurn } from '@/types/conversationTurn';
import { SheetAction } from '@/types/sheet-actions';

class MemoryStorage implements Storage {
  private store = new Map<string, string>();

  get length(): number {
    return this.store.size;
  }

  clear(): void {
    this.store.clear();
  }

  getItem(key: string): string | null {
    return this.store.get(key) ?? null;
  }

  key(index: number): string | null {
    return Array.from(this.store.keys())[index] ?? null;
  }

  removeItem(key: string): void {
    this.store.delete(key);
  }

  setItem(key: string, value: string): void {
    this.store.set(key, value);
  }
}

describe('chatSessionStorage', () => {
  beforeEach(() => {
    vi.stubGlobal('localStorage', new MemoryStorage());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('round-trips sessions and revives processing turns as complete', () => {
    const session = createChatSession();
    const withTurn = appendUserTurnToSession(session, 'Find 2290', 'turn_1');
    const store = {
      activeSessionId: withTurn.id,
      sessions: [withTurn],
    };

    saveChatSessions('Budget_2026.xlsx', store);
    const loaded = loadChatSessions('Budget_2026.xlsx');

    expect(loaded?.activeSessionId).toBe(withTurn.id);
    expect(loaded?.sessions).toHaveLength(1);
    expect(loaded?.sessions[0]?.title).toBe('Find 2290');
    expect(loaded?.sessions[0]?.turns[0]?.userMessage).toBe('Find 2290');
    expect(loaded?.sessions[0]?.turns[0]?.phase).toBe('complete');
    expect(loaded?.sessions[0]?.history[0]?.content).toBe('Find 2290');
  });

  it('returns null when workbook key is missing or storage is empty', () => {
    expect(loadChatSessions('')).toBeNull();
    expect(loadChatSessions('missing-workbook')).toBeNull();
  });

  it('round-trips assistantMode per workbook', () => {
    const store = {
      activeSessionId: null,
      sessions: [],
      assistantMode: 'plan' as const,
    };

    saveChatSessions('Budget_2026.xlsx', store);
    const loaded = loadChatSessions('Budget_2026.xlsx');

    expect(loaded?.assistantMode).toBe('plan');
  });

  it('round-trips gstPurchasePortalPreference per workbook', () => {
    const store = {
      activeSessionId: null,
      sessions: [],
      gstPurchasePortalPreference: 'gstr2b_only' as const,
    };

    saveChatSessions('Budget_2026.xlsx', store);
    const loaded = loadChatSessions('Budget_2026.xlsx');

    expect(loaded?.gstPurchasePortalPreference).toBe('gstr2b_only');
  });

  describe('attachment imports', () => {
    const ROW = [46123, 46123, 'UPIOUT/123456789012/payee@upi', 'S12345678', 1912, '', 632.04, 'p1 l22', ''];

    function importBlock(overrides: Partial<ActionBlock> = {}): ActionBlock {
      return {
        id: 'actions_import_1',
        type: 'actions',
        actions: [
          { type: 'CREATE_SHEET', sheetName: 'Bank Statement', name: 'Bank Statement' },
          { type: 'WRITE_TABLE', sheetName: 'Bank Statement', headers: ['Date', 'Debit'], rows: [ROW, ROW] },
          { type: 'FREEZE_PANES', sheetName: 'Bank Statement', freezeRows: 1 },
        ] as SheetAction[],
        explanation: 'Create sheet "Bank Statement" with these transactions?',
        proposalStatus: 'pending',
        importSheetName: 'Bank Statement',
        ...overrides,
      };
    }

    function roundTrip(block: ActionBlock): { saved: string; loaded: ActionBlock } {
      const session = appendUserTurnToSession(createChatSession(), 'Import bank statement: s.pdf', 'turn_1');
      const turn: ConversationTurn = { ...session.turns[0], phase: 'complete', blocks: [block] };
      saveChatSessions('Book.xlsx', { activeSessionId: session.id, sessions: [{ ...session, turns: [turn] }] });
      const saved = localStorage.getItem('cellix:chat-sessions:Book.xlsx') ?? '';
      const loaded = loadChatSessions('Book.xlsx')?.sessions[0]?.turns[0]?.blocks[0] as ActionBlock;
      return { saved, loaded };
    }

    it('does not save the imported transactions', () => {
      const { saved } = roundTrip(importBlock());
      expect(saved).not.toContain('UPIOUT/123456789012');
      expect(saved).not.toContain('S12345678');
      expect(saved).not.toContain('632.04');
      // The card itself, and the rest of its actions, are still saved.
      expect(saved).toContain('actions_import_1');
      expect(saved).toContain('FREEZE_PANES');
    });

    it('leaves the live block untouched when saving', () => {
      const block = importBlock();
      roundTrip(block);
      expect((block.actions[1] as { rows: unknown[][] }).rows).toHaveLength(2);
    });

    it('closes a pending import on reload, since it can no longer be applied', () => {
      const { loaded } = roundTrip(importBlock());
      expect(loaded.proposalStatus).toBe('rejected');
      expect(loaded.userFacingSummary?.supportingDetail).toBe(
        'This import was not accepted before the pane was closed. Attach the file again to import it.',
      );
      expect((loaded.actions[1] as { rows: unknown[][] }).rows).toEqual([]);
    });

    it('keeps an accepted import shown as accepted', () => {
      const { loaded } = roundTrip(importBlock({ proposalStatus: 'accepted' }));
      expect(loaded.proposalStatus).toBe('accepted');
    });

    it('saves the rows of every other kind of card in full', () => {
      const { saved, loaded } = roundTrip(importBlock({ importSheetName: undefined }));
      expect(saved).toContain('UPIOUT/123456789012');
      expect(loaded.proposalStatus).toBe('pending');
      expect((loaded.actions[1] as { rows: unknown[][] }).rows).toHaveLength(2);
    });

    // TASKS.md #404 — "Reverted" is read from the block, so it has to come back with it:
    // a pane reloaded after a Revert must not show the change as applied and offer it again.
    it('keeps a reverted change reverted across a reload of the pane', () => {
      const { loaded } = roundTrip(
        importBlock({ importSheetName: undefined, proposalStatus: 'accepted', changeSetId: 'cs_1', reverted: true }),
      );
      expect(loaded.proposalStatus).toBe('accepted');
      expect(loaded.reverted).toBe(true);
    });
  });
});
