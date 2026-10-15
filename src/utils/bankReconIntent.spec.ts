import { describe, expect, it } from 'vitest';
import {
  detectBankReconIntent,
  detectBankReconLayout,
  detectBankReconRelayoutTarget,
  isAttachedReconPrompt,
  isBankReconPrompt,
  isBankReconRelayoutPrompt,
  isCasualBankRecon,
} from './bankReconIntent';
import { isGstReconPrompt } from './gstReconIntent';

describe('bankReconIntent', () => {
  it.each([
    'reconcile bank',
    'bank reconciliation for April 2025',
    'prepare BRS',
    'match bank statement with ledger',
    'reconcile my bank book vs statement',
  ])('detects "%s"', (msg) => {
    expect(isBankReconPrompt(msg)).toBe(true);
  });

  it.each([
    'reconcile purchase register with GSTR-2B',
    'bank reconciliation of GST ITC',
    'what is the bank holiday list',
    'sum column D',
  ])('ignores "%s"', (msg) => {
    expect(isBankReconPrompt(msg)).toBe(false);
  });

  it('never overlaps with the GST detector', () => {
    for (const msg of ['reconcile bank', 'prepare BRS', 'reconcile purchase register with 2B']) {
      expect(isBankReconPrompt(msg) && isGstReconPrompt(msg)).toBe(false);
    }
  });

  it('is casual unless an account number or full report is requested', () => {
    const casual = detectBankReconIntent('reconcile bank for April 2025')!;
    expect(isCasualBankRecon('reconcile bank for April 2025', casual)).toBe(true);
    const gated = detectBankReconIntent('reconcile bank account no 50100123456789')!;
    expect(gated.extractedAccountNumber).toBe('50100123456789');
    expect(isCasualBankRecon('reconcile bank account no 50100123456789', gated)).toBe(false);
  });

  it('picks layouts from phrasing', () => {
    expect(detectBankReconLayout('reconcile bank, show entries to post')).toBe('bank_flat');
    expect(detectBankReconLayout('reconcile bank, cheques not cleared')).toBe('books_flat');
    expect(detectBankReconLayout('reconcile bank and show what matched')).toBe('matched');
    expect(detectBankReconLayout('reconcile bank')).toBe('categorized');
  });

  it('recognizes relayout follow-ups', () => {
    expect(isBankReconRelayoutPrompt('change the layout of Bank Recon Summary to a single table')).toBe(true);
    expect(detectBankReconRelayoutTarget('change the Bank Recon layout to a single table')).toBe('books_flat');
    expect(detectBankReconRelayoutTarget('switch Bank Recon layout to bank only')).toBe('bank_flat');
    expect(isBankReconRelayoutPrompt('change the layout of Missed vs GSTR-2B to a single table')).toBe(false);
  });
});

describe('timing-only view', () => {
  it('narrows only for cleared/presented/transit wording', async () => {
    const { wantsTimingItemsOnly } = await import('./bankReconIntent');
    expect(wantsTimingItemsOnly('reconcile bank, cheques not cleared')).toBe(true);
    expect(wantsTimingItemsOnly('reconcile bank, outstanding cheques')).toBe(true);
    expect(wantsTimingItemsOnly('reconcile bank, single table')).toBe(false);
  });
});

describe('isAttachedReconPrompt (a file is attached)', () => {
  it.each([
    'reconcile bank ledger against this file',
    'reconcile bank with this file',
    'reconcile this',
    'match my ledger with the attached statement',
    'compare my books to this',
    'tally this against my cash book',
    'do the BRS for September',
    'recon',
    'check it against my bank ledger',
    'Reconcile @[Bank Ledger] against the attached CSV',
  ])('treats "%s" as reconcile', (msg) => {
    expect(isAttachedReconPrompt(msg)).toBe(true);
  });

  it.each(['import this bank statement', 'read this statement into a sheet', 'what is the closing balance', '', 'reconcile with GSTR-2B', 'match purchase register with 2B'])(
    'leaves "%s" to the import / GST flow',
    (msg) => {
      expect(isAttachedReconPrompt(msg)).toBe(false);
    },
  );
});
