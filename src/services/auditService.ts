import {
  getAuditApplyEndpoint,
  getAuditExportEndpoint,
  getAuditHistoryEndpoint,
  getAuditPreviewLocalEndpoint,
  getAuditRevertEndpoint,
  getAuditStatsEndpoint,
} from '@/lib/apiConfig';
import { CellChange, ChangeSetSummary } from '@/types/changeSet';
import { SheetAction } from '@/types/sheet-actions';
import type { CreatedConditionalFormatId, CreatedChartId } from '@/engine/actionEngine';

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

  const response = await fetch(url, { credentials: 'include', ...init, headers });
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

export async function markChangeSetApplied(
  changeSetId: string,
  createdConditionalFormatIds?: CreatedConditionalFormatId[],
  createdChartIds?: CreatedChartId[],
  // TASKS.md #93 — the real before/after diff for a SORT_RANGE, read directly
  // off Excel by the frontend. The backend's shadow-based diff skips sparse
  // ranges, so without this Revert has nothing to undo for those sorts.
  sortedRangeChanges?: CellChange[],
): Promise<ChangeSetSummary> {
  const body: Record<string, unknown> = {};
  // Fastify rejects Content-Type: application/json with an empty body (400) —
  // '{}' is the pre-existing default; TASKS.md #40/#15 add optional keys so the
  // backend can patch the real Excel-assigned id into a pending CONDITIONAL_FORMAT
  // or CREATE_CHART structuralOp before marking the change set applied.
  if (createdConditionalFormatIds && createdConditionalFormatIds.length > 0) {
    body.createdConditionalFormatIds = createdConditionalFormatIds;
  }
  if (createdChartIds && createdChartIds.length > 0) {
    body.createdChartIds = createdChartIds;
  }
  if (sortedRangeChanges && sortedRangeChanges.length > 0) {
    body.sortedRangeChanges = sortedRangeChanges;
  }
  const result = await auditFetch<{ changeSet: ChangeSetSummary }>(
    getAuditApplyEndpoint(changeSetId),
    {
      method: 'POST',
      body: JSON.stringify(body),
    },
  );
  return result.changeSet;
}

export interface LocalChangeSetResult {
  changeSetId: string;
  irreversibleActionTypes: string[];
  changes: CellChange[];
}

/**
 * Registers a change set for actions the client already resolved and will
 * apply itself with no LLM call at all (the local sheet-action fast lanes —
 * copy/rename/delete/create-empty sheet). Without this, those actions never
 * got a `changeSetId` at all, so Revert never appeared for them — not
 * specific to any one action type, a gap in the fast lane itself.
 *
 * Deliberately sends a MINIMAL synthetic context (sheet names only, no real
 * cell values) rather than doing a real Office.js data read here — that
 * would defeat the entire point of the fast lane being instant. This is
 * safe specifically for the action types the fast lane produces: ADD_SHEET /
 * RENAME_SHEET / DELETE_SHEET are pure structural operations (their revert
 * path never needs cell data), and ADD_SHEET{copyFrom}'s destination cells
 * are excluded from the diff server-side regardless of source completeness
 * (change-set.service.ts's copiedDestSheets filter, TASKS.md #248).
 *
 * Never throws on the caller's behalf in the way the callers actually use
 * it — see dispatchLocalSheetActions, which treats a failure here as "no
 * revert available this time" rather than blocking the local action.
 */
export async function previewLocalChangeSet(
  conversationId: string,
  prompt: string,
  actions: SheetAction[],
  sheetNames: string[],
  activeSheetName: string,
): Promise<LocalChangeSetResult> {
  const context = {
    activeSheetName: activeSheetName || sheetNames[0] || 'Sheet1',
    sheets: sheetNames.map((name) => ({
      name,
      usedRange: 'A1:A1',
      rowCount: 1,
      columnCount: 1,
      values: [[null]],
      formulas: [['']],
      numberFormats: [['General']],
      structure: 'unknown' as const,
    })),
    namedRanges: [],
    tables: [],
  };

  const result = await auditFetch<{
    changeSet: { changeSetId: string; irreversibleActionTypes: string[]; changes: CellChange[] };
  }>(getAuditPreviewLocalEndpoint(), {
    method: 'POST',
    body: JSON.stringify({ conversationId, prompt, context, actions }),
  });
  return {
    changeSetId: result.changeSet.changeSetId,
    irreversibleActionTypes: result.changeSet.irreversibleActionTypes ?? [],
    changes: result.changeSet.changes ?? [],
  };
}

export async function revertChangeSet(
  changeSetId: string,
): Promise<{ changeSet: ChangeSetSummary; inverseActions: SheetAction[] }> {
  return auditFetch(getAuditRevertEndpoint(changeSetId), {
    method: 'POST',
    body: '{}',
  });
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

  const response = await fetch(url, { headers, credentials: 'include' });
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
