import { getConversationByIdEndpoint, getConversationListEndpoint } from '@/lib/apiConfig';
import { StoredConversation } from '@/utils/rehydrateConversation';

/**
 * One row of server-backed chat history (TASKS.md #171/#172).
 *
 * Summary only — the list deliberately carries no message bodies, so opening the
 * history panel stays cheap no matter how much history a user has. Full content
 * arrives from `fetchConversationById` when a row is actually clicked.
 */
export interface ConversationSummary {
  conversationId: string;
  workbookId?: string;
  title: string;
  firstMessage: string;
  lastMessage: string;
  messageCount: number;
  status: string;
  updatedAt: string | null;
}

export interface ConversationHistoryPage {
  conversations: ConversationSummary[];
  nextCursor: string | null;
}

/**
 * Same thin per-service fetch wrapper convention as checkpointService.ts /
 * auditService.ts — each service owns its own so error semantics don't get
 * coupled across modules.
 */
async function historyFetch<T>(url: string, init?: RequestInit): Promise<T> {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    ...(init?.headers as Record<string, string> | undefined),
  };
  if (url.includes('.ngrok-free.app')) {
    headers['ngrok-skip-browser-warning'] = 'true';
  }

  const response = await fetch(url, { credentials: 'include', ...init, headers });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Conversation history API ${response.status}: ${text || response.statusText}`);
  }

  // 204 No Content (delete) has no body to parse.
  if (response.status === 204) {
    return undefined as T;
  }

  const body: unknown = await response.json();
  // The list route is @SkipEnvelope()'d, but unwrap defensively so this keeps
  // working if that decorator is ever dropped.
  if (body && typeof body === 'object' && 'data' in body) {
    return (body as { data: T }).data;
  }
  return body as T;
}

export async function fetchConversationHistory(params?: {
  limit?: number;
  cursor?: string;
  workbookId?: string;
}): Promise<ConversationHistoryPage> {
  const page = await historyFetch<ConversationHistoryPage>(getConversationListEndpoint(params));
  return {
    conversations: page?.conversations ?? [],
    nextCursor: page?.nextCursor ?? null,
  };
}

export async function fetchConversationById(
  conversationId: string,
): Promise<StoredConversation> {
  return historyFetch<StoredConversation>(getConversationByIdEndpoint(conversationId));
}

/** Rename a conversation (TASKS.md #177). Returns the saved (trimmed) title. */
export async function renameConversation(
  conversationId: string,
  title: string,
): Promise<{ conversationId: string; title: string }> {
  return historyFetch(getConversationByIdEndpoint(conversationId), {
    method: 'PATCH',
    body: JSON.stringify({ title }),
  });
}

/** Delete a conversation (TASKS.md #177). Hard delete — not reversible. */
export async function deleteConversation(conversationId: string): Promise<void> {
  await historyFetch<void>(getConversationByIdEndpoint(conversationId), { method: 'DELETE' });
}
