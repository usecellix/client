/**
 * Client API for POST /gst/reconcile
 */

import { getApiBaseUrl } from '@/lib/apiConfig';
import { SheetAction } from '@/types/sheet-actions';
import { UserFacingSummary } from '@/utils/userFacingResponse';

export function getGstReconcileEndpoint(): string {
  return `${getApiBaseUrl()}/gst/reconcile`;
}

export type GstReconType =
  | 'PR_VS_GSTR2B'
  | 'PR_VS_GSTR2A'
  | 'IMS_VS_PR'
  | 'GSTR3B_VS_GSTR2B'
  | 'SALES_VS_GSTR1';

export interface SheetPayload {
  sheet_name: string;
  headers_row?: number;
  data: unknown[][];
  column_mapping?: Record<string, string | number>;
  file_type?: string;
}

export interface GstReconcileRequest {
  reconciliation_type: GstReconType;
  purchase_register?: SheetPayload;
  portal_file?: SheetPayload;
  ims_data?: SheetPayload | null;
  settings?: {
    amount_tolerance_abs?: number;
    amount_tolerance_pct?: number;
    invoice_fuzzy_threshold?: number;
    date_tolerance_days?: number;
    detect_rcm?: boolean;
    use_ims_data?: boolean;
  };
  period?: string;
  gstin?: string;
  conversation_id?: string;
  output_sheet_name?: string;
}

export interface GstReconSummary {
  total_pr_rows: number;
  total_portal_rows: number;
  total_ims_rows: number;
  exact_matched: number;
  partial_matched: number;
  credit_notes: number;
  pr_only: number;
  portal_only: number;
  ims_rejected: number;
  ims_pending: number;
  ims_auto_accept: number;
  rcm_flagged: number;
  itc_matched: number;
  itc_at_risk: number;
  rcm_payable: number;
  itc_ims_rejected: number;
  itc_ims_pending: number;
}

export interface GstReconcileResponse {
  job_id: string;
  status: string;
  reconciliation_type: string;
  summary: GstReconSummary;
  rows: Array<{
    status: string;
    invoice_number?: string | null;
    gstin?: string | null;
    itc_amount?: number;
    difference?: string | null;
  }>;
  actions: SheetAction[];
  confidence: number;
  exceptions: unknown[];
  output_sheet_name: string;
  audit_log_id: string | null;
}

function formatInr(n: number): string {
  return new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0,
  }).format(n || 0);
}

export function buildGstReconUserFacingSummary(
  result: GstReconcileResponse,
  prName: string,
  portalName: string,
): UserFacingSummary {
  const s = result.summary;
  return {
    contextLine: `${prName} ↔ ${portalName}`,
    headline: `GST reconciliation ready — Accept to create sheet "${result.output_sheet_name}".`,
    bullets: [
      `Exact matched: ${s.exact_matched} · Partial: ${s.partial_matched} · Credit notes: ${s.credit_notes}`,
      `Unmatched in books (PR only): ${s.pr_only} · Unmatched on portal only: ${s.portal_only}`,
      `ITC matched: ${formatInr(s.itc_matched)} · ITC at risk: ${formatInr(s.itc_at_risk)}`,
      ...(s.rcm_flagged
        ? [`RCM flags: ${s.rcm_flagged} (payable ~ ${formatInr(s.rcm_payable)})`]
        : []),
    ],
    supportingDetail: `Job ${result.job_id} · CA to verify before filing. Nothing is written until you Accept.`,
  };
}

export function buildGstReconAnswerText(
  result: GstReconcileResponse,
  prName: string,
  portalName: string,
): string {
  const s = result.summary;
  const unmatchedSample = result.rows
    .filter((r) => r.status === 'PR_ONLY' || r.status === 'PORTAL_ONLY')
    .slice(0, 8)
    .map((r) => `• ${r.status}: ${r.invoice_number ?? '?'} (${r.gstin ?? 'no GSTIN'})`)
    .join('\n');

  return [
    `Finished matching **${prName}** against **${portalName}**.`,
    '',
    `**Summary**`,
    `- Exact matched: ${s.exact_matched}`,
    `- Partial matches: ${s.partial_matched}`,
    `- Credit notes: ${s.credit_notes}`,
    `- In books only (not on portal): ${s.pr_only}`,
    `- On portal only (not in books): ${s.portal_only}`,
    `- ITC matched: ${formatInr(s.itc_matched)}`,
    `- ITC at risk: ${formatInr(s.itc_at_risk)}`,
    unmatchedSample ? `\n**Sample unmatched lines**\n${unmatchedSample}` : '',
    '',
    `Review the card below and **Accept** to create sheet **${result.output_sheet_name}** with matched and unmatched entries, or **Reject** to discard.`,
  ]
    .filter(Boolean)
    .join('\n');
}

async function reconcileFetch<T>(url: string, body: unknown): Promise<T> {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  };
  if (url.includes('.ngrok-free.app')) {
    headers['ngrok-skip-browser-warning'] = 'true';
  }

  const response = await fetch(url, {
    method: 'POST',
    credentials: 'include',
    headers,
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`GST reconcile failed (${response.status}): ${text || response.statusText}`);
  }

  const bodyJson: unknown = await response.json();
  if (bodyJson && typeof bodyJson === 'object' && 'data' in bodyJson) {
    return (bodyJson as { data: T }).data;
  }
  return bodyJson as T;
}

export async function runGstReconcile(
  request: GstReconcileRequest,
): Promise<GstReconcileResponse> {
  // Map SALES_VS_GSTR1 to PR_VS_GSTR2B-shaped payload (server treats register vs portal generically)
  const payload: GstReconcileRequest = { ...request };
  if (payload.reconciliation_type === 'SALES_VS_GSTR1') {
    payload.reconciliation_type = 'PR_VS_GSTR2B';
    if (payload.portal_file) {
      payload.portal_file = { ...payload.portal_file, file_type: 'GSTR2B' };
    }
  }

  return reconcileFetch<GstReconcileResponse>(getGstReconcileEndpoint(), payload);
}
