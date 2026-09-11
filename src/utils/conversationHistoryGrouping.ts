import { ConversationSummary } from '@/services/conversationHistoryService';

export interface ConversationHistoryGroup {
  label: string;
  conversations: ConversationSummary[];
}

const DAY_MS = 24 * 60 * 60 * 1000;

/** Midnight local time, so "Yesterday" means the calendar day — not 24h ago. */
function startOfDay(date: Date): number {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

function bucketLabel(updatedAt: string | null, now: Date): string {
  if (!updatedAt) return 'Older';
  const timestamp = new Date(updatedAt).getTime();
  if (Number.isNaN(timestamp)) return 'Older';

  const today = startOfDay(now);
  if (timestamp >= today) return 'Today';
  if (timestamp >= today - DAY_MS) return 'Yesterday';
  if (timestamp >= today - 7 * DAY_MS) return 'Previous 7 days';
  if (timestamp >= today - 30 * DAY_MS) return 'Previous 30 days';
  return 'Older';
}

const ORDER = ['Today', 'Yesterday', 'Previous 7 days', 'Previous 30 days', 'Older'];

/**
 * Bucket history rows into ChatGPT-style recency sections (TASKS.md #172).
 *
 * Input order is preserved within each group — the server already sorted by
 * `updatedAt` descending, and re-sorting here would just risk disagreeing with
 * the cursor pagination that depends on that same ordering.
 *
 * Empty groups are dropped, so a user with only old chats doesn't scroll past
 * four empty headings to reach them.
 */
export function groupConversationsByRecency(
  conversations: ConversationSummary[],
  now: Date = new Date(),
): ConversationHistoryGroup[] {
  const buckets = new Map<string, ConversationSummary[]>();

  for (const conversation of conversations) {
    const label = bucketLabel(conversation.updatedAt, now);
    const existing = buckets.get(label);
    if (existing) {
      existing.push(conversation);
    } else {
      buckets.set(label, [conversation]);
    }
  }

  return ORDER.filter((label) => buckets.get(label)?.length).map((label) => ({
    label,
    conversations: buckets.get(label)!,
  }));
}

/**
 * Drop duplicate conversationIds while keeping first-seen order.
 *
 * Needed when appending a page: a conversation updated between two page fetches
 * can legitimately appear in both, and rendering it twice would give it two rows
 * with the same React key.
 */
export function dedupeConversations(
  conversations: ConversationSummary[],
): ConversationSummary[] {
  const seen = new Set<string>();
  const result: ConversationSummary[] = [];
  for (const conversation of conversations) {
    if (seen.has(conversation.conversationId)) continue;
    seen.add(conversation.conversationId);
    result.push(conversation);
  }
  return result;
}
