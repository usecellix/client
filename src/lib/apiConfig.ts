const DEFAULT_API_BASE_URL = '/api';
const STREAM_PATH = '/excel-ai/process';
const CONVERSATION_PATH = '/excel-ai/conversation';
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

export function getCompareEndpoint(): string {
  return `${getApiBaseUrl()}${COMPARE_PATH}`;
}

