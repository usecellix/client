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

describe('layout-aware chat summaries for the two flat layouts', () => {
  it('books_flat: short summary states the count and direction, not the full reason breakdown', () => {
    const result = makeResult(
      makeSummary({ pr_only: 30, gstin_mismatch_count: 2, portal_only: 0 }),
    );
    result.output_sheet_name = 'Missed vs GSTR-2B — Books Only';

    const text = buildGstReconAnswerText(result, 'Purchase register', 'GSTR-2B', 'books_flat');
    expect(text).toContain('32');
    expect(text).toContain('exist in your books but not in GSTR-2B');
    expect(text).toContain('Accept to create **Missed vs GSTR-2B — Books Only**');
    // Never the categorized layout's full reason breakdown.
    expect(text).not.toContain('by reason:');
    expect(text).not.toContain('grouped by reason');

    const summary = buildGstReconUserFacingSummary(result, 'Purchase register', 'GSTR-2B', 'books_flat');
    expect(summary.bullets?.some((b) => b.includes('32') && b.includes('your books'))).toBe(true);
  });

  it('portal_flat: short summary uses the mirrored (reverse) direction wording', () => {
    const result = makeResult(makeSummary({ pr_only: 0, gstin_mismatch_count: 0, portal_only: 15 }));
    result.output_sheet_name = 'Missed vs GSTR-2B — Portal Only';

    const text = buildGstReconAnswerText(result, 'Purchase register', 'GSTR-2B', 'portal_flat');
    expect(text).toContain('15');
    expect(text).toContain('exist in GSTR-2B but not in your books');
    expect(text).toContain('Accept to create **Missed vs GSTR-2B — Portal Only**');
    expect(text).not.toContain('by reason:');

    const summary = buildGstReconUserFacingSummary(result, 'Purchase register', 'GSTR-2B', 'portal_flat');
    expect(summary.bullets?.some((b) => b.includes('15'))).toBe(true);
  });

  it('portal_flat creates a sheet based on portal_only alone, even when there are zero books-side issues (the gap a books-only hasIssues check would miss)', () => {
    const result = makeResult(makeSummary({ pr_only: 0, gstin_mismatch_count: 0, portal_only: 4 }));

    const text = buildGstReconAnswerText(result, 'Purchase register', 'GSTR-2B', 'portal_flat');
    expect(text).not.toContain("No rows found");
    expect(text).toContain('Create a sheet with these rows?');

    const summary = buildGstReconUserFacingSummary(result, 'Purchase register', 'GSTR-2B', 'portal_flat');
    expect(summary.headline).toBe('Create a sheet with these rows?');
  });

  it('portal_flat reports no rows and no sheet when there are zero portal-only rows, regardless of books-side issues', () => {
    const result = makeResult(makeSummary({ pr_only: 12, gstin_mismatch_count: 3, portal_only: 0 }));

    const text = buildGstReconAnswerText(result, 'Purchase register', 'GSTR-2B', 'portal_flat');
    expect(text).toContain('No rows found in GSTR-2B that are missing from your books.');
    expect(text).toContain('No sheet will be created.');

    const summary = buildGstReconUserFacingSummary(result, 'Purchase register', 'GSTR-2B', 'portal_flat');
    expect(summary.headline).toContain('No rows found');
  });

  it('categorized (default, layout omitted) still gets the full reason-breakdown text — unchanged', () => {
    const result = makeResult(makeSummary({ pr_only: 5, gstin_mismatch_count: 0 }));
    const text = buildGstReconAnswerText(result, 'Purchase register', 'B2B');
    expect(text).toContain('Not matched — **5** rows, by reason:');
    expect(text).toContain('Create a sheet with full details, grouped by reason?');
  });
});

/**
 * The casual-flow answer text no longer embeds an inline "Sample missed invoices" bullet
 * list (capped at 8, plain text, nothing clickable) — the full missed-row list is now a
 * separate clickable card (`GstReconMissedRowsCard`, built from `extractMissedRows` in
 * gstReconChat.ts) that jumps to the exact books-sheet cell per row. This describe block
 * now only confirms the answer TEXT itself never regresses back to leaking a bare
 * "• ()"-shaped artifact for a blank-field row — the actual per-row list/blank-field
 * fallback coverage lives in gstReconChat.spec.ts against `extractMissedRows` instead.
 */
describe('casual-flow answer text never leaks a bare "• ()" bullet artifact for blank-field rows', () => {
  it('produces no bullet-list line at all for the missed rows (moved to the missed-rows card)', () => {
    const result = makeResult(makeSummary({ pr_only: 1, gstin_mismatch_count: 0 }));
    result.rows = [
      {
        status: 'PR_ONLY',
        invoice_number: '',
        gstin: '',
      },
    ];

    const text = buildGstReconAnswerText(result, 'Purchase register', 'B2B');
    expect(text).not.toContain('• ()');
    expect(text).not.toContain('Sample missed invoices');
  });
});

describe('RCM and Amended counts surface in the casual-flow chat summary (v2.0 spec gap fix)', () => {
  it('reports RCM rows and treats them as an issue even when nothing else is unmatched', () => {
    const result = makeResult(
      makeSummary({ pr_only: 0, gstin_mismatch_count: 0, rcm_flagged: 3 }),
    );

    const text = buildGstReconAnswerText(result, 'Purchase register', 'GSTR-2B');
    expect(text).toContain('Possible RCM (reverse charge)');
    expect(text).toContain('3');
    expect(text).toContain('Create a sheet with full details');

    const summary = buildGstReconUserFacingSummary(result, 'Purchase register', 'GSTR-2B');
    expect(summary.headline).toContain('Create a sheet');
    expect(summary.bullets?.some((b) => b.includes('Possible RCM') && b.includes('3'))).toBe(
      true,
    );
  });

  it('reports amended-invoice rows and treats them as an issue even when nothing else is unmatched', () => {
    const result = makeResult(
      makeSummary({ pr_only: 0, gstin_mismatch_count: 0, amended_count: 2 }),
    );

    const text = buildGstReconAnswerText(result, 'Purchase register', 'GSTR-2B');
    expect(text).toContain('Amended invoices');
    expect(text).toContain('verify against the original');

    const summary = buildGstReconUserFacingSummary(result, 'Purchase register', 'GSTR-2B');
    expect(summary.headline).toContain('Create a sheet');
    expect(
      summary.bullets?.some((b) => b.includes('Amended invoices') && b.includes('2')),
    ).toBe(true);
  });

  it('omits both bullets and reports "no missed rows" when RCM and amended counts are zero, same as before', () => {
    const result = makeResult(
      makeSummary({ pr_only: 0, gstin_mismatch_count: 0, rcm_flagged: 0, amended_count: 0 }),
    );

    const text = buildGstReconAnswerText(result, 'Purchase register', 'GSTR-2B');
    expect(text).not.toContain('Possible RCM');
    expect(text).not.toContain('Amended invoices');
    expect(text).toContain("No missed rows — I won't create a sheet.");
  });
});
