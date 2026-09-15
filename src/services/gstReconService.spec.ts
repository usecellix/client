import { describe, it, expect } from 'vitest';
import {
  buildGstReconAnswerText,
  buildGstReconUserFacingSummary,
  GstReconcileResponse,
  GstReconSummary,
} from './gstReconService';

function makeSummary(overrides: Partial<GstReconSummary> = {}): GstReconSummary {
  return {
    total_pr_rows: 33,
    total_portal_rows: 35,
    total_ims_rows: 0,
    exact_matched: 24,
    partial_matched: 0,
    credit_notes: 0,
    pr_only: 0,
    portal_only: 2,
    ims_rejected: 0,
    ims_pending: 0,
    ims_auto_accept: 0,
    rcm_flagged: 0,
    itc_matched: 0,
    itc_at_risk: 0,
    rcm_payable: 0,
    itc_ims_rejected: 0,
    itc_ims_pending: 0,
    matched_exact: 0,
    matched_fallback: 24,
    mismatch_blank_gstin: 0,
    mismatch_blank_taxable_value: 0,
    mismatch_ambiguous_rate_slab: 0,
    mismatch_amount: 0,
    mismatch_date: 0,
    mismatch_genuinely_missing: 0,
    gstin_mismatch_count: 0,
    ...overrides,
  };
}

function makeResult(summary: GstReconSummary): GstReconcileResponse {
  return {
    job_id: 'j1',
    status: 'complete',
    reconciliation_type: 'PR_VS_GSTR2B',
    summary,
    rows: [],
    actions: [],
    confidence: 1,
    exceptions: [],
    output_sheet_name: 'Missed vs GSTR-2B',
    audit_log_id: null,
    missed_books_only: true,
  };
}

/**
 * Regression coverage: the "Not matched — N rows, by reason:" header was gated on
 * pr_only alone, so a run where GSTIN mismatch was the ONLY non-zero reason (pr_only=0,
 * gstin_mismatch_count=9 — a real May 2024 run) silently dropped the header even though
 * the sheet and total count were both correct. The header must key off total not-matched
 * (pr_only + gstin_mismatch_count), not any single bucket, so phrasing stays consistent
 * regardless of which reason(s) happen to be non-zero.
 */
describe('"Not matched" header keys off total not-matched, not any single reason bucket', () => {
  it('renders when GSTIN mismatch is the only non-zero reason (pr_only=0)', () => {
    const result = makeResult(makeSummary({ pr_only: 0, gstin_mismatch_count: 9 }));

    const text = buildGstReconAnswerText(result, 'Purchase register', 'B2B');
    expect(text).toContain('Not matched — **9** rows, by reason:');

    const summary = buildGstReconUserFacingSummary(result, 'Purchase register', 'B2B');
    expect(summary.bullets?.some((b) => b.includes('Not matched — 9 rows, by reason:'))).toBe(true);
  });

  it('shows the combined total when both pr_only and gstin_mismatch_count are non-zero', () => {
    const result = makeResult(makeSummary({ pr_only: 5, gstin_mismatch_count: 3 }));
    const text = buildGstReconAnswerText(result, 'Purchase register', 'B2B');
    expect(text).toContain('Not matched — **8** rows, by reason:');
  });

  it('still works when pr_only is the only non-zero reason (gstin_mismatch_count=0)', () => {
    const result = makeResult(makeSummary({ pr_only: 5, gstin_mismatch_count: 0 }));
    const text = buildGstReconAnswerText(result, 'Purchase register', 'B2B');
    expect(text).toContain('Not matched — **5** rows, by reason:');
  });

  it('omits the header and reports "no missed rows" when both are zero', () => {
    const result = makeResult(makeSummary({ pr_only: 0, gstin_mismatch_count: 0 }));
    const text = buildGstReconAnswerText(result, 'Purchase register', 'B2B');
    expect(text).not.toContain('Not matched');
    expect(text).toContain("No missed rows — I won't create a sheet.");

    const summary = buildGstReconUserFacingSummary(result, 'Purchase register', 'B2B');
    expect(summary.headline).toContain('No missed rows');
  });
});
