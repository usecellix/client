import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createChatSession } from '@/types/chatSession';
import { appendUserTurnToSession } from '@/utils/sessionContinuity';
import { loadChatSessions, saveChatSessions } from '@/utils/chatSessionStorage';

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
});
