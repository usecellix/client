import {
  getAuditApplyEndpoint,
  getAuditExportEndpoint,
  getAuditHistoryEndpoint,
  getAuditRevertEndpoint,
  getAuditStatsEndpoint,
} from '@/lib/apiConfig';
import { ChangeSetSummary } from '@/types/changeSet';
import { SheetAction } from '@/types/sheet-actions';

export interface AuditStats {
  totalCost: number;
  totalTokens: number;
  totalCalls: number;
  successRate: number;
  avgLatencyMs: number;
  byTier: Record<string, { calls: number; cost: number; tokens: number }>;
}

async function auditFetch<T>(url: string, init?: RequestInit): Promise<T> {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    ...(init?.headers as Record<string, string> | undefined),
  };
  if (url.includes('.ngrok-free.app')) {
    headers['ngrok-skip-browser-warning'] = 'true';
  }

  const response = await fetch(url, { ...init, headers });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Audit API ${response.status}: ${text || response.statusText}`);
  }

  const body: unknown = await response.json();
  if (body && typeof body === 'object' && 'data' in body) {
    return (body as { data: T }).data;
  }
  return body as T;
}

export async function markChangeSetApplied(changeSetId: string): Promise<ChangeSetSummary> {
  const result = await auditFetch<{ changeSet: ChangeSetSummary }>(
    getAuditApplyEndpoint(changeSetId),
    { method: 'POST' },
  );
  return result.changeSet;
}

export async function revertChangeSet(
  changeSetId: string,
): Promise<{ changeSet: ChangeSetSummary; inverseActions: SheetAction[] }> {
  return auditFetch(getAuditRevertEndpoint(changeSetId), { method: 'POST' });
}

export async function fetchChangeSetHistory(
  conversationId: string,
): Promise<ChangeSetSummary[]> {
  const result = await auditFetch<{ changeSets: ChangeSetSummary[] }>(
    getAuditHistoryEndpoint(conversationId),
  );
  return result.changeSets ?? [];
}

export async function fetchAuditStats(from?: string, to?: string): Promise<AuditStats> {
  return auditFetch<AuditStats>(getAuditStatsEndpoint(from, to));
}

export async function downloadAuditExport(
  format: 'json' | 'csv',
  from?: string,
  to?: string,
): Promise<void> {
  const url = getAuditExportEndpoint(format, from, to);
  const headers: Record<string, string> = {};
  if (url.includes('.ngrok-free.app')) {
    headers['ngrok-skip-browser-warning'] = 'true';
  }

  const response = await fetch(url, { headers });
  if (!response.ok) {
    throw new Error(`Export failed (${response.status})`);
  }

  const blob = await response.blob();
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = objectUrl;
  anchor.download = `cellix-audit.${format === 'csv' ? 'csv' : 'json'}`;
  anchor.click();
  URL.revokeObjectURL(objectUrl);
}
