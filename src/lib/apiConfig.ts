const DEFAULT_API_BASE_URL = '/api';
const STREAM_PATH = '/excel-ai/process';
const CONVERSATION_PATH = '/excel-ai/conversation';
const CONVERSATIONS_PATH = '/excel-ai/conversations';
const COMPARE_PATH = '/sheets/compare';

function normalizeBaseUrl(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) {
    throw new Error('VITE_API_BASE_URL is set but empty. Provide a valid URL like /api');
  }
  return trimmed.replace(/\/+$/, '');
}

export function getApiBaseUrl(): string {
  const envValue = (import.meta as any)?.env?.VITE_API_BASE_URL as string | undefined;

  if (typeof envValue === 'string') {
    return normalizeBaseUrl(envValue);
  }

  if ((import.meta as any)?.env?.DEV) {
    // eslint-disable-next-line no-console
    console.warn(
      `[Cellix] VITE_API_BASE_URL is not set. Falling back to ${DEFAULT_API_BASE_URL}. ` +
        `Create frontend/.env with VITE_API_BASE_URL=/api to use the dev proxy.`,
    );
  }

  return DEFAULT_API_BASE_URL;
}

export function getStreamEndpoint(): string {
  return `${getApiBaseUrl()}${STREAM_PATH}`;
}

export function getConversationEndpoint(): string {
  return `${getApiBaseUrl()}${CONVERSATION_PATH}`;
}

export function getConversationByIdEndpoint(conversationId: string): string {
  return `${getApiBaseUrl()}${CONVERSATION_PATH}/${encodeURIComponent(conversationId)}`;
}

/** Server-backed chat history for the signed-in user (TASKS.md #171/#172). */
export function getConversationListEndpoint(params?: {
  limit?: number;
  cursor?: string;
  workbookId?: string;
}): string {
  const query = new URLSearchParams();
  if (params?.limit !== undefined) query.set('limit', String(params.limit));
  if (params?.cursor) query.set('cursor', params.cursor);
  if (params?.workbookId) query.set('workbookId', params.workbookId);
  const suffix = query.toString();
  return `${getApiBaseUrl()}${CONVERSATIONS_PATH}${suffix ? `?${suffix}` : ''}`;
}

export function getToolResultEndpoint(): string {
  return `${getApiBaseUrl()}${CONVERSATION_PATH}/tool-result`;
}

export function getCompareEndpoint(): string {
  return `${getApiBaseUrl()}${COMPARE_PATH}`;
}

export function getAuditApplyEndpoint(changeSetId: string): string {
  return `${getApiBaseUrl()}/audit/apply/${encodeURIComponent(changeSetId)}`;
}

export function getAuditRevertEndpoint(changeSetId: string): string {
  return `${getApiBaseUrl()}/audit/revert/${encodeURIComponent(changeSetId)}`;
}

export function getAuditHistoryEndpoint(conversationId: string): string {
  return `${getApiBaseUrl()}/audit/history/${encodeURIComponent(conversationId)}`;
}

export function getAuditStatsEndpoint(from?: string, to?: string): string {
  const params = new URLSearchParams();
  if (from) params.set('from', from);
  if (to) params.set('to', to);
  const query = params.toString();
  return `${getApiBaseUrl()}/audit/stats${query ? `?${query}` : ''}`;
}

export function getAuditExportEndpoint(format: 'json' | 'csv', from?: string, to?: string): string {
  const params = new URLSearchParams({ format });
  if (from) params.set('from', from);
  if (to) params.set('to', to);
  return `${getApiBaseUrl()}/audit/export?${params.toString()}`;
}

export function getCheckpointCreateEndpoint(): string {
  return `${getApiBaseUrl()}/audit/checkpoint`;
}

export function getCheckpointListEndpoint(workbookId: string): string {
  return `${getApiBaseUrl()}/audit/checkpoint/${encodeURIComponent(workbookId)}`;
}

export function getCheckpointRestoreEndpoint(checkpointId: string): string {
  return `${getApiBaseUrl()}/audit/checkpoint/restore/${encodeURIComponent(checkpointId)}`;
}

