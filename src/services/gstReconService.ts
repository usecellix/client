/**
 * Client API for POST /gst/reconcile
 */

import { getApiBaseUrl } from '@/lib/apiConfig';
import { SheetAction } from '@/types/sheet-actions';
import { UserFacingSummary } from '@/utils/userFacingResponse';

export function getGstReconcileEndpoint(): string {
  return `${getApiBaseUrl()}/gst/reconcile`;
}

export function getGstReconAuditOutcomeEndpoint(auditLogId: string): string {
  return `${getApiBaseUrl()}/gst/reconcile/audit/${encodeURIComponent(auditLogId)}/outcome`;
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
  books_register?: SheetPayload;
  portal_file?: SheetPayload;
  portal_file_2a?: SheetPayload | null;
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
  financial_year?: string;
  client_name?: string;
  /** Required client taxpayer GSTIN */
  client_gstin?: string;
  /** @deprecated prefer client_gstin */
  gstin?: string;
  conversation_id?: string;
  output_sheet_name?: string;
  missed_books_only?: boolean;
  /** Output layout for the missed_books_only sheet — see GstReconLayout. Ignored otherwise. */
  layout?: 'categorized' | 'books_flat' | 'portal_flat';
  /** Inclusive date range (yyyy-mm-dd) extracted from the prompt — filters books/portal rows before matching. */
  period_start?: string;
  period_end?: string;
  /** Human-readable period label for chat text, e.g. "April 2024". */
  period_label?: string;
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
  cross_gstin_exception_count?: number;
  /** Matched breakdown by pass. */
  matched_exact?: number;
  matched_fallback?: number;
  /** Unmatched-in-books breakdown by reason. */
  mismatch_blank_gstin?: number;
  /** Blank-GSTIN rows resolved against a unique portal_only candidate — subset of mismatch_blank_gstin's original count, no longer double-counted in portal_only. */
  mismatch_blank_gstin_likely_matched?: number;
  mismatch_blank_taxable_value?: number;
  mismatch_ambiguous_rate_slab?: number;
  mismatch_amount?: number;
  mismatch_date?: number;
  mismatch_genuinely_missing?: number;
  /** Same vendor (PAN), same invoice, booked under a different GSTIN registration on each side — its own bucket. */
  gstin_mismatch_count?: number;
}

export interface GstReconcileResponse {
  job_id: string;
  status: string;
  reconciliation_type: string;
  actionType?: string;
  client_gstin?: string | null;
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
  cross_gstin_exception_count?: number;
  missed_books_only?: boolean;
  portal_label?: string;
  books_label?: string;
  /** Present when a period filter was applied to this run. */
  period_applied?: { label: string; start: string; end: string };
  /** Present instead of running the reconciliation when the stated period matched zero rows in books or portal. */
  period_zero_message?: string;
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
  const isSales = result.reconciliation_type === 'SALES_VS_GSTR1';
  const portalLabel = result.portal_label ?? portalName;
  const booksLabel = result.books_label ?? (isSales ? 'Sales Register' : 'Purchase Register');
  const periodSuffix = result.period_applied ? ` (${result.period_applied.label})` : '';
  const missed = s.pr_only;
  const gstinMismatchCount = s.gstin_mismatch_count ?? 0;
  const hasIssues = missed > 0 || gstinMismatchCount > 0;
  const totalNotMatched = missed + gstinMismatchCount;

  if (result.missed_books_only) {
    const matchedExact = s.matched_exact ?? s.exact_matched;
    const matchedFallback = s.matched_fallback ?? 0;
    const reasonBullets = [
      s.mismatch_blank_gstin ? `Blank GSTIN in the register: ${s.mismatch_blank_gstin}` : null,
      s.mismatch_blank_gstin_likely_matched
        ? `Blank GSTIN, but likely matched to a portal invoice (confirm & fill in): ${s.mismatch_blank_gstin_likely_matched}`
        : null,
      s.mismatch_blank_taxable_value
        ? `Taxable value blank (no rate column populated): ${s.mismatch_blank_taxable_value}`
        : null,
      s.mismatch_ambiguous_rate_slab
        ? `Ambiguous rate slab, needs CA review: ${s.mismatch_ambiguous_rate_slab}`
        : null,
      s.mismatch_amount
        ? `Amount differs from the portal record for the same GSTIN + date: ${s.mismatch_amount}`
        : null,
      s.mismatch_date
        ? `Date differs from the portal record for the same GSTIN + amount: ${s.mismatch_date}`
        : null,
      s.mismatch_genuinely_missing
        ? `Genuinely missing from ${portalLabel}: ${s.mismatch_genuinely_missing}`
        : null,
    ].filter((b): b is string => Boolean(b));

    return {
      contextLine: `${prName} ↔ ${portalLabel}${periodSuffix}`,
      headline: hasIssues
        ? `Create a sheet with the missed rows, grouped by reason?`
        : `No missed rows — every ${booksLabel} line matched ${portalLabel}.`,
      bullets: hasIssues
        ? [
            `Matched: ${matchedExact} exact · ${matchedFallback} fallback, no invoice no. · ${s.partial_matched} partial · ${s.credit_notes} credit notes`,
            ...(totalNotMatched > 0
              ? [
                  `Not matched — ${totalNotMatched} row${totalNotMatched === 1 ? '' : 's'}, by reason:`,
                  ...reasonBullets,
                ]
              : []),
            ...(gstinMismatchCount > 0
              ? [
                  `Possible GSTIN mismatch, same vendor under a different registration: ${gstinMismatchCount}`,
                ]
              : []),
            `Accept to create sheet "${result.output_sheet_name}" with the full breakdown`,
          ]
        : [
            `Checked ${s.total_pr_rows} books rows against ${s.total_portal_rows} portal rows`,
            `Matched: ${matchedExact} exact · ${matchedFallback} fallback, no invoice no. · ${s.partial_matched} partial`,
          ],
      supportingDetail: hasIssues
        ? `Nothing is written until you Accept. Reject discards this sheet.`
        : `No sheet will be created.`,
    };
  }

  const taxMatchedLabel = isSales ? 'Tax liability matched' : 'ITC matched';
  const taxAtRiskLabel = isSales ? 'Tax liability at risk' : 'ITC at risk';
  const crossCount =
    result.cross_gstin_exception_count ?? s.cross_gstin_exception_count ?? 0;
  const matchedFallback = s.matched_fallback ?? 0;
  return {
    contextLine: `${prName} ↔ ${portalName}`,
    headline: `GST reconciliation ready — Accept to create sheet "${result.output_sheet_name}".`,
    bullets: [
      `Matched: ${s.matched_exact ?? s.exact_matched} exact${matchedFallback ? ` · ${matchedFallback} fallback` : ''} · ${s.partial_matched} partial · ${s.credit_notes} credit notes`,
      `In books only: ${s.pr_only} (${s.mismatch_blank_gstin ?? 0} blank GSTIN · ${s.mismatch_blank_gstin_likely_matched ?? 0} blank GSTIN likely matched · ${s.mismatch_blank_taxable_value ?? 0} blank value · ${s.mismatch_ambiguous_rate_slab ?? 0} ambiguous rate slab · ${s.mismatch_amount ?? 0} amount mismatch · ${s.mismatch_date ?? 0} date mismatch · ${s.mismatch_genuinely_missing ?? 0} genuinely missing) · On portal only: ${s.portal_only}`,
      `${taxMatchedLabel}: ${formatInr(s.itc_matched)} · ${taxAtRiskLabel}: ${formatInr(s.itc_at_risk)}`,
      ...(crossCount > 0
        ? [`Cross-GSTIN exceptions: ${crossCount} (flagged, not matched)`]
        : []),
      ...(s.rcm_flagged
        ? [`RCM flags: ${s.rcm_flagged} (payable ~ ${formatInr(s.rcm_payable)})`]
        : []),
    ],
    supportingDetail: `Job ${result.job_id}${
      result.client_gstin ? ` · GSTIN ${result.client_gstin}` : ''
    } · CA to verify before filing. Nothing is written until you Accept.`,
  };
}

export function buildGstReconAnswerText(
  result: GstReconcileResponse,
  prName: string,
  portalName: string,
): string {
  const s = result.summary;
  const isSales = result.reconciliation_type === 'SALES_VS_GSTR1';
  const portalLabel = result.portal_label ?? portalName;
  const booksLabel = result.books_label ?? (isSales ? 'Sales Register' : 'Purchase Register');
  const periodNote = result.period_applied ? ` for **${result.period_applied.label}**` : '';
  const missed = s.pr_only;
  const gstinMismatchCount = s.gstin_mismatch_count ?? 0;
  const hasIssues = missed > 0 || gstinMismatchCount > 0;
  const totalNotMatched = missed + gstinMismatchCount;
  const missedSample = result.rows
    .filter((r) => r.status === 'PR_ONLY')
    .slice(0, 8)
    .map((r) => `• ${r.invoice_number ?? '?'} (${r.gstin ?? 'no GSTIN'})`)
    .join('\n');

  if (result.missed_books_only) {
    const matchedExact = s.matched_exact ?? s.exact_matched;
    const matchedFallback = s.matched_fallback ?? 0;
    const matchedTotal = matchedExact + matchedFallback + s.partial_matched + s.credit_notes;

    if (!hasIssues) {
      return [
        `I compared **${prName}** with **${portalLabel}**${periodNote}.`,
        '',
        `All **${s.total_pr_rows}** ${booksLabel} rows were found on the portal (${matchedExact} exact · ${matchedFallback} fallback, no invoice no. · ${s.partial_matched} partial).`,
        '',
        `No missed rows — I won't create a sheet.`,
      ].join('\n');
    }

    const reasonLines = [
      s.mismatch_blank_gstin
        ? `  • ${s.mismatch_blank_gstin} — GSTIN blank in the register`
        : null,
      s.mismatch_blank_gstin_likely_matched
        ? `  • ${s.mismatch_blank_gstin_likely_matched} — GSTIN blank, but likely matched to a portal invoice on date + amount (confirm and fill in the GSTIN)`
        : null,
      s.mismatch_blank_taxable_value
        ? `  • ${s.mismatch_blank_taxable_value} — Taxable value blank (no rate column populated)`
        : null,
      s.mismatch_ambiguous_rate_slab
        ? `  • ${s.mismatch_ambiguous_rate_slab} — Ambiguous rate slab (needs CA review)`
        : null,
      s.mismatch_amount
        ? `  • ${s.mismatch_amount} — Amount differs from the portal record for the same GSTIN + date`
        : null,
      s.mismatch_date
        ? `  • ${s.mismatch_date} — Date differs from the portal record for the same GSTIN + amount`
        : null,
      s.mismatch_genuinely_missing
        ? `  • ${s.mismatch_genuinely_missing} — Genuinely missing from ${portalLabel} (not filed by supplier, or filed elsewhere)`
        : null,
    ].filter((l): l is string => Boolean(l));

    return [
      `Compared **${s.total_pr_rows}** rows in **${prName}** against ${portalLabel}${periodNote}.`,
      '',
      `Matched: **${matchedTotal}** (${matchedExact} exact · ${matchedFallback} fallback, no invoice no. · ${s.partial_matched} partial · ${s.credit_notes} credit/debit note)`,
      '',
      totalNotMatched > 0
        ? `Not matched — **${totalNotMatched}** row${totalNotMatched === 1 ? '' : 's'}, by reason:`
        : '',
      ...reasonLines,
      gstinMismatchCount > 0
        ? `\n**Possible GSTIN mismatch** — ${gstinMismatchCount} invoice${gstinMismatchCount === 1 ? '' : 's'} where the same vendor (same PAN) appears to be booked under a different GSTIN registration on each side.`
        : '',
      missedSample ? `\n**Sample missed invoices**\n${missedSample}` : '',
      '',
      `**Create a sheet with full details, grouped by reason?**`,
      `Accept to create **${result.output_sheet_name}**, or Reject to discard.`,
    ]
      .filter(Boolean)
      .join('\n');
  }

  const taxMatchedLabel = isSales ? 'Tax liability matched' : 'ITC matched';
  const taxAtRiskLabel = isSales ? 'Tax liability at risk' : 'ITC at risk';
  const crossCount =
    result.cross_gstin_exception_count ?? s.cross_gstin_exception_count ?? 0;
  const unmatchedSample = result.rows
    .filter(
      (r) =>
        r.status === 'PR_ONLY' ||
        r.status === 'PORTAL_ONLY' ||
        r.status === 'CROSS_GSTIN' ||
        r.status === 'GSTIN_MISMATCH',
    )
    .slice(0, 8)
    .map((r) => `• ${r.status}: ${r.invoice_number ?? '?'} (${r.gstin ?? 'no GSTIN'})`)
    .join('\n');

  const fullMatchedFallback = s.matched_fallback ?? 0;
  return [
    `Finished matching **${prName}** against **${portalName}**.`,
    '',
    `**Summary**`,
    `- Exact matched: ${s.matched_exact ?? s.exact_matched}`,
    fullMatchedFallback ? `- Matched (fallback, no invoice no.): ${fullMatchedFallback}` : '',
    `- Partial matches: ${s.partial_matched}`,
    `- Credit notes: ${s.credit_notes}`,
    `- In books only (not on portal): ${s.pr_only}, by reason:`,
    `  • ${s.mismatch_blank_gstin ?? 0} — blank GSTIN`,
    `  • ${s.mismatch_blank_gstin_likely_matched ?? 0} — blank GSTIN, likely matched to a portal invoice (confirm)`,
    `  • ${s.mismatch_blank_taxable_value ?? 0} — blank taxable value`,
    `  • ${s.mismatch_ambiguous_rate_slab ?? 0} — ambiguous rate slab (needs CA review)`,
    `  • ${s.mismatch_amount ?? 0} — amount mismatch`,
    `  • ${s.mismatch_date ?? 0} — date mismatch`,
    `  • ${s.mismatch_genuinely_missing ?? 0} — genuinely missing`,
    (s.gstin_mismatch_count ?? 0) > 0
      ? `- Possible GSTIN mismatch (same vendor, different registration): ${s.gstin_mismatch_count}`
      : '',
    `- On portal only (not in books): ${s.portal_only}`,
    `- ${taxMatchedLabel}: ${formatInr(s.itc_matched)}`,
    `- ${taxAtRiskLabel}: ${formatInr(s.itc_at_risk)}`,
    crossCount > 0 ? `- Cross-GSTIN exceptions: ${crossCount}` : '',
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
  return reconcileFetch<GstReconcileResponse>(getGstReconcileEndpoint(), request);
}

export async function reportGstReconAuditOutcome(
  auditLogId: string,
  outcome: 'applied' | 'rejected',
): Promise<void> {
  if (!auditLogId) return;
  try {
    await reconcileFetch(getGstReconAuditOutcomeEndpoint(auditLogId), { outcome });
  } catch {
    // Non-blocking — Excel apply already succeeded/failed independently
  }
}
