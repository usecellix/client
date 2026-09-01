import { describe, expect, it } from 'vitest';
import {
  dedupeConversations,
  groupConversationsByRecency,
} from '@/utils/conversationHistoryGrouping';
import { ConversationSummary } from '@/services/conversationHistoryService';

function summary(overrides: Partial<ConversationSummary> = {}): ConversationSummary {
  return {
    conversationId: 'conv_1',
    title: 'Sum the totals',
    firstMessage: 'Sum the totals',
    lastMessage: 'Done.',
    messageCount: 2,
    status: 'active',
    updatedAt: '2026-09-01T09:00:00.000Z',
    ...overrides,
  };
}

/** Fixed "now" so these assertions don't drift with the wall clock. */
const NOW = new Date('2026-09-01T18:00:00.000Z');

describe('groupConversationsByRecency (TASKS.md #172)', () => {
  it('buckets by calendar recency', () => {
    const groups = groupConversationsByRecency(
      [
        summary({ conversationId: 'today', updatedAt: NOW.toISOString() }),
        summary({
          conversationId: 'yesterday',
          updatedAt: new Date(NOW.getTime() - 24 * 3600_000).toISOString(),
        }),
        summary({
          conversationId: 'week',
          updatedAt: new Date(NOW.getTime() - 4 * 24 * 3600_000).toISOString(),
        }),
        summary({
          conversationId: 'month',
          updatedAt: new Date(NOW.getTime() - 20 * 24 * 3600_000).toISOString(),
        }),
        summary({
          conversationId: 'ancient',
          updatedAt: new Date(NOW.getTime() - 200 * 24 * 3600_000).toISOString(),
        }),
      ],
      NOW,
    );

    expect(groups.map((group) => group.label)).toEqual([
      'Today',
      'Yesterday',
      'Previous 7 days',
      'Previous 30 days',
      'Older',
    ]);
  });

  it('omits empty groups instead of rendering bare headings', () => {
    const groups = groupConversationsByRecency(
      [summary({ updatedAt: NOW.toISOString() })],
      NOW,
    );

    expect(groups).toHaveLength(1);
    expect(groups[0].label).toBe('Today');
  });

  it('preserves the server ordering within a group', () => {
    const groups = groupConversationsByRecency(
      [
        summary({ conversationId: 'newer', updatedAt: '2026-09-01T12:00:00.000Z' }),
        summary({ conversationId: 'older', updatedAt: '2026-09-01T08:00:00.000Z' }),
      ],
      NOW,
    );

    expect(groups[0].conversations.map((c) => c.conversationId)).toEqual(['newer', 'older']);
  });

  it('files a missing or unparseable timestamp under Older rather than dropping the row', () => {
    const groups = groupConversationsByRecency(
      [
        summary({ conversationId: 'no-date', updatedAt: null }),
        summary({ conversationId: 'bad-date', updatedAt: 'not-a-date' }),
      ],
      NOW,
    );

    expect(groups).toEqual([
      expect.objectContaining({ label: 'Older' }),
    ]);
    expect(groups[0].conversations).toHaveLength(2);
  });

  it('returns nothing for an empty history', () => {
    expect(groupConversationsByRecency([], NOW)).toEqual([]);
  });
});

describe('dedupeConversations', () => {
  it('drops a conversation that appears in two fetched pages', () => {
    const result = dedupeConversations([
      summary({ conversationId: 'a' }),
      summary({ conversationId: 'b' }),
      summary({ conversationId: 'a' }),
    ]);

    expect(result.map((c) => c.conversationId)).toEqual(['a', 'b']);
  });

  it('keeps first-seen order', () => {
    const result = dedupeConversations([
      summary({ conversationId: 'b' }),
      summary({ conversationId: 'a' }),
    ]);

    expect(result.map((c) => c.conversationId)).toEqual(['b', 'a']);
  });
});
