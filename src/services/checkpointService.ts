import {
  getCheckpointCreateEndpoint,
  getCheckpointListEndpoint,
  getCheckpointRestoreEndpoint,
} from '@/lib/apiConfig';
import { CheckpointSummary, RestoreResult } from '@/types/checkpoint';

/**
 * TASKS.md #31. Same fetch-wrapper convention as auditService.ts's auditFetch —
 * kept as a private duplicate rather than a shared import to avoid coupling this
 * module's error semantics to auditService's, matching the existing pattern where
 * each service file owns its own thin fetch wrapper.
 */
async function checkpointFetch<T>(url: string, init?: RequestInit): Promise<T> {
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
    throw new Error(`Checkpoint API ${response.status}: ${text || response.statusText}`);
  }

  const body: unknown = await response.json();
  if (body && typeof body === 'object' && 'data' in body) {
    return (body as { data: T }).data;
  }
  return body as T;
}

export async function fetchCheckpoints(workbookId: string): Promise<CheckpointSummary[]> {
  const result = await checkpointFetch<{ checkpoints: CheckpointSummary[] }>(
    getCheckpointListEndpoint(workbookId),
  );
  return result.checkpoints ?? [];
}

export async function createCheckpoint(
  workbookId: string,
  conversationId: string,
  label?: string,
): Promise<CheckpointSummary> {
  const result = await checkpointFetch<{ checkpoint: CheckpointSummary }>(
    getCheckpointCreateEndpoint(),
    {
      method: 'POST',
      body: JSON.stringify({ workbookId, conversationId, label }),
    },
  );
  return result.checkpoint;
}

export async function restoreCheckpoint(checkpointId: string): Promise<RestoreResult> {
  return checkpointFetch<RestoreResult>(getCheckpointRestoreEndpoint(checkpointId), {
    method: 'POST',
    body: '{}',
  });
}
