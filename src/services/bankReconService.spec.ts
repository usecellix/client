import { describe, expect, it } from 'vitest';
import { buildBankReconAnswerText, BankReconcileResponse } from './bankReconService';

const base = (over: Partial<BankReconcileResponse['summary']> = {}): BankReconcileResponse =>
  ({
    job_id: 'j',
    status: 'complete',
    summary: {
      books_rows: 57, bank_rows: 65, matched: 43, matched_group: 2, matched_books_rows: 48, reversal_pairs: 0,
      books_only: 10, bank_only: 20, low_confidence_rows: 0,
      books_reasons: { genuinely_missing: 2 }, bank_reasons: { bank_charges_not_booked: 8 },
      ...over,
    },
    rows: [],
    actions: [],
    brs: { periodEnd: '2026-09-30', booksBalance: 1, bankBalance: 1, difference: 0, tiedOut: true, notes: [] },
    output_sheet_name: 'Bank Recon Summary',
    layout: 'categorized',
    layout_cache_id: null,
    period_applied: { label: 'Sep 2026', start: '2026-09-01', end: '2026-09-30' },
    carry_forward_items: 3,
    validation: { errors: [], warnings: [], notes: [] },
  }) as unknown as BankReconcileResponse;

describe('buildBankReconAnswerText wording', () => {
  it('uses real plurals (entries, matches, items) — not "entrys" / "matchs"', () => {
    const text = buildBankReconAnswerText(base(), 'Bank Ledger', 'x.csv', 'categorized');
    expect(text).toContain('10 ledger entries not in the bank');
    expect(text).toContain('20 bank entries not in your books');
    expect(text).toContain('2 group matches');
    expect(text).toContain('3 items carried forward');
    expect(text).not.toMatch(/entrys|matchs/);
  });

  it('uses the singular for exactly one', () => {
    const text = buildBankReconAnswerText(base({ books_only: 1, bank_only: 1, matched_group: 1 }), 'Bank Ledger', 'x.csv', 'categorized');
    expect(text).toContain('1 ledger entry not in the bank');
    expect(text).toContain('1 bank entry not in your books');
    expect(text).toContain('1 group match ');
  });
});

describe('buildBankReconAnswerText layout', () => {
  const attached = () =>
    ({
      ...base(),
      bank_source: { kind: 'attachment', file_name: 'bank_statement_sep2026_B_semicolon_cp1252.csv', file_type: 'csv', transactions: 65 },
    }) as BankReconcileResponse;

  it('keeps every bullet together (the attached-file note comes before them) so the list reads as one block', () => {
    const r = attached();
    const lines = buildBankReconAnswerText(r, 'Bank Ledger', r.bank_source!.file_name, 'categorized').split('\n');
    const bullets = lines.map((l, i) => (l.startsWith('• ') ? i : -1)).filter((i) => i >= 0);
    expect(bullets.length).toBeGreaterThanOrEqual(4);
    expect(bullets[bullets.length - 1] - bullets[0]).toBe(bullets.length - 1); // no non-bullet line in between
    const note = lines.findIndex((l) => l.startsWith('Statement read directly'));
    expect(note).toBeGreaterThan(-1);
    expect(note).toBeLessThan(bullets[0]);
  });

  it('types in a few seconds at the single app-wide speed', async () => {
    const r = attached();
    const text = buildBankReconAnswerText(r, 'Bank Ledger', r.bank_source!.file_name, 'categorized');
    const { TYPING_CHARS_PER_TICK, TYPING_INTERVAL_MS } = await import('@/components/ConversationPanel/AnswerReveal');
    expect(text.length).toBeGreaterThan(400);
    expect((Math.ceil(text.length / TYPING_CHARS_PER_TICK) * TYPING_INTERVAL_MS) / 1000).toBeLessThan(5);
  });
});
