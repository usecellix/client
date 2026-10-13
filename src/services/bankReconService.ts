/**
 * Client API for POST /bank/reconcile (+ relayout) and result → chat-text helpers.
 */

import { getApiBaseUrl } from '@/lib/apiConfig';
import { SheetAction } from '@/types/sheet-actions';
import { BankReconLayout } from '@/utils/bankReconIntent';
import { UserFacingSummary } from '@/utils/userFacingResponse';

export function getBankReconcileEndpoint(): string {
  return `${getApiBaseUrl()}/bank/reconcile`;
}

export function getBankRelayoutEndpoint(): string {
  return `${getApiBaseUrl()}/bank/reconcile/relayout`;
}

export interface BankSheetPayload {
  sheet_name: string;
  data: unknown[][];
  headers_row?: number;
  file_type?: 'XLSX' | 'CSV';
}

export interface BankReconcileRequest {
  books_ledger: BankSheetPayload;
  bank_statement: BankSheetPayload;
  /** Previous period's BRS sheet; its uncleared items are carried forward. */
  prior_brs?: BankSheetPayload;
  period_start?: string;
  period_end?: string;
  period_label?: string;
  casual?: boolean;
  layout?: BankReconLayout;
  books_filter?: 'timing';
  account_number?: string;
  client_name?: string;
  books_opening_balance?: number;
  bank_opening_balance?: number;
  conversation_id?: string;
}

export interface BankReconRow {
  side: 'BOOKS' | 'BANK';
  sheet_name: string;
  row: number;
  date: string;
  direction: 'IN' | 'OUT';
  amount: number;
  narration: string;
  reference: string;
  reason: string;
  reason_label: string;
  explanation: string;
}

export interface BankReconBrs {
  periodEnd: string;
  booksBalance: number | null;
  bankBalance: number | null;
  difference: number | null;
  tiedOut: boolean;
  notes: string[];
}

export interface BankReconSummary {
  books_rows: number;
  bank_rows: number;
  matched: number;
  matched_group: number;
  matched_books_rows?: number;
  reversal_pairs: number;
  books_only: number;
  bank_only: number;
  low_confidence_rows: number;
  books_reasons: Record<string, number>;
  bank_reasons: Record<string, number>;
}

export interface BankReconcileResponse {
  job_id: string;
  status: 'complete';
  summary: BankReconSummary;
  rows: BankReconRow[];
  actions: SheetAction[];
  brs: BankReconBrs | null;
  output_sheet_name: string;
  layout: BankReconLayout;
  layout_cache_id: string | null;
  period_applied: { label: string; start: string; end: string };
  carry_forward_items?: number;
  period_zero_message?: string;
  validation: {
    errors: Array<{ code: string; message: string }>;
    warnings: Array<{ code: string; message: string }> | string[];
    notes: string[];
  };
}

export interface BankRelayoutResponse {
  actions: SheetAction[];
  output_sheet_name: string;
  layout: BankReconLayout;
}

function inr(n: number | null | undefined): string {
  return new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 2 }).format(n ?? 0);
}

const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? '' : 's'}`;

export function buildBankReconAnswerText(
  result: BankReconcileResponse,
  booksName: string,
  bankName: string,
  layout: BankReconLayout,
): string {
  const s = result.summary;
  const period = result.period_applied.label;
  const matched = s.matched_books_rows ?? s.matched + s.matched_group;
  const lines: string[] = [
    `**Bank reconciliation — ${period}** (${booksName} ↔ ${bankName})`,
    `Matched ${matched} of ${s.books_rows} ledger entries against ${s.bank_rows} statement rows.`,
  ];
  if (s.books_only) {
    const top = Object.entries(s.books_reasons)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([k, v]) => `${v} ${k.replace(/_/g, ' ')}`)
      .join(', ');
    lines.push(`• ${plural(s.books_only, 'ledger entry')} not in the bank (${top}).`);
  }
  if (s.bank_only) {
    const top = Object.entries(s.bank_reasons)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([k, v]) => `${v} ${k.replace(/_/g, ' ')}`)
      .join(', ');
    lines.push(`• ${plural(s.bank_only, 'bank entry')} not in your books (${top}) — these need to be posted.`);
  }
  if (!s.books_only && !s.bank_only) lines.push('• Everything in the period reconciles.');
  if (result.carry_forward_items) lines.push(`• ${plural(result.carry_forward_items, 'item')} carried forward from the prior BRS.`);
  if (s.reversal_pairs) lines.push(`• ${plural(s.reversal_pairs, 'reversal / bounce pair')} netted off.`);
  if (s.matched_group) lines.push(`• ${plural(s.matched_group, 'group match')} (one entry covering several) — please confirm.`);
  if (s.low_confidence_rows) lines.push(`• ${plural(s.low_confidence_rows, 'row')} read with lower confidence — listed under “Please verify”.`);

  const brs = result.brs;
  if (brs) {
    if (brs.tiedOut) lines.push(`BRS ties out to ₹0.00 (books ${inr(brs.booksBalance)} / bank ${inr(brs.bankBalance)}).`);
    else if (brs.difference !== null) lines.push(`⚠ BRS does **not** tie out — residual difference ${inr(brs.difference)}. Treat this as unreconciled.`);
    else lines.push('BRS balances could not be read from the sheets, so the statement is not tied out — the reconciling items above are still complete.');
  }
  const errs = result.validation.errors;
  if (errs.length) lines.push(`Data checks: ${errs.map((e) => e.message).join(' ')}`);
  if (result.validation.notes?.length) lines.push(...result.validation.notes);

  const label =
    layout === 'books_flat' ? 'a books-only table' : layout === 'bank_flat' ? 'a bank-only (entries to post) table' : layout === 'matched' ? 'the matched list' : 'the reconciliation summary';
  lines.push(`**Accept** to create sheet “${result.output_sheet_name}” with ${label}, or **Reject** to discard.`);
  return lines.join('\n');
}

export function buildBankReconUserFacingSummary(
  result: BankReconcileResponse,
  booksName: string,
  bankName: string,
): UserFacingSummary {
  const s = result.summary;
  const brs = result.brs;
  return {
    contextLine: `${booksName} ↔ ${bankName} (${result.period_applied.label})`,
    headline: 'Create the bank reconciliation sheet?',
    bullets: [
      `${s.matched_books_rows ?? s.matched + s.matched_group} ledger entries matched · ${s.books_only} in books only · ${s.bank_only} in bank only`,
      brs
        ? brs.tiedOut
          ? 'BRS ties out to 0.00'
          : brs.difference !== null
            ? `BRS difference ${inr(brs.difference)}`
            : 'BRS balances not found'
        : '',
      `Accept to create sheet "${result.output_sheet_name}"`,
    ].filter(Boolean),
    supportingDetail: 'Nothing is written until you Accept. Reject discards this sheet.',
  };
}

async function postJson<T>(url: string, body: unknown): Promise<T> {
  const headers: Record<string, string> = { 'Content-Type': 'application/json' };
  if (url.includes('.ngrok-free.app')) headers['ngrok-skip-browser-warning'] = 'true';
  const response = await fetch(url, { method: 'POST', credentials: 'include', headers, body: JSON.stringify(body) });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    let message = text || response.statusText;
    try {
      const parsed = JSON.parse(text) as { message?: string | string[] };
      if (parsed?.message) message = Array.isArray(parsed.message) ? parsed.message.join('; ') : parsed.message;
    } catch {
      /* keep raw text */
    }
    throw new Error(message);
  }
  const json: unknown = await response.json();
  if (json && typeof json === 'object' && 'data' in json) return (json as { data: T }).data;
  return json as T;
}

export function runBankReconcile(request: BankReconcileRequest): Promise<BankReconcileResponse> {
  return postJson<BankReconcileResponse>(getBankReconcileEndpoint(), request);
}

export function runBankRelayout(request: {
  layout_cache_id: string;
  layout: BankReconLayout;
}): Promise<BankRelayoutResponse> {
  return postJson<BankRelayoutResponse>(getBankRelayoutEndpoint(), request);
}
