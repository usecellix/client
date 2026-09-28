import { describe, expect, it } from 'vitest';
import {
  detectGstReconIntent,
  detectGstReconLayout,
  isCasualMissedRowsRecon,
  isGstReconPrompt,
} from './gstReconIntent';

describe('gstReconIntent', () => {
  it('detects purchase vs GSTR-2B prompts', () => {
    expect(detectGstReconIntent('reconcile my GST purchases with GSTR-2B')?.type).toBe(
      'PR_VS_GSTR2B',
    );
    expect(isGstReconPrompt('Match purchase register to 2B')).toBe(true);
  });

  it('detects 2A and IMS', () => {
    expect(detectGstReconIntent('reconcile PR with GSTR-2A')?.type).toBe('PR_VS_GSTR2A');
    expect(detectGstReconIntent('IMS vs purchase register reconciliation')?.type).toBe(
      'IMS_VS_PR',
    );
  });

  it('ignores non-GST chat', () => {
    expect(detectGstReconIntent('create a new sheet called Summary')).toBeNull();
    expect(detectGstReconIntent('what is the total in column B?')).toBeNull();
  });

  it('detects casual purchase and sales register recon', () => {
    const purchase = detectGstReconIntent('Reconcile purchase register');
    expect(purchase?.type).toBe('PR_VS_GSTR2B');
    expect(isCasualMissedRowsRecon('Reconcile purchase register', purchase!)).toBe(true);

    const sales = detectGstReconIntent('Reconcile sales register');
    expect(sales?.type).toBe('SALES_VS_GSTR1');
    expect(isCasualMissedRowsRecon('Reconcile sales register', sales!)).toBe(true);
  });

  it('does not treat GSTIN-gated prompts as casual', () => {
    const intent = detectGstReconIntent(
      'Reconcile purchase register for GSTIN 27ABCDE1234F1Z5',
    );
    expect(intent).not.toBeNull();
    expect(isCasualMissedRowsRecon('Reconcile purchase register for GSTIN 27ABCDE1234F1Z5', intent!)).toBe(
      false,
    );
  });
});

describe('detectGstReconLayout', () => {
  it('defaults to categorized when no layout phrasing is present', () => {
    expect(detectGstReconLayout('Reconcile purchase register')).toBe('categorized');
    expect(detectGstReconLayout('')).toBe('categorized');
  });

  it('selects books_flat on explicit "just show missed rows" / "simple layout" phrasing', () => {
    expect(detectGstReconLayout('just show missed rows')).toBe('books_flat');
    expect(detectGstReconLayout('Reconcile purchase register, just show me the missed rows')).toBe(
      'books_flat',
    );
    expect(detectGstReconLayout('give me a simple layout please')).toBe('books_flat');
  });

  it('selects books_flat for "rows in my books not in the portal" — books side named first', () => {
    expect(detectGstReconLayout('rows in my books not in the portal')).toBe('books_flat');
    expect(detectGstReconLayout('show invoices in my purchase register not in GSTR-2B')).toBe(
      'books_flat',
    );
  });

  it('selects portal_flat for "rows in the portal not in my books" — portal side named first (reverse direction)', () => {
    expect(detectGstReconLayout('rows in the portal not in my books')).toBe('portal_flat');
    expect(detectGstReconLayout('show invoices in GSTR-2B not in my register')).toBe('portal_flat');
  });

  it('selects portal_flat for "invoices the portal has that I haven\'t recorded"', () => {
    expect(
      detectGstReconLayout("invoices the portal has that I haven't recorded"),
    ).toBe('portal_flat');
  });

  it('never guesses a layout from ambiguous phrasing like "what\'s missing in GSTR-2B"', () => {
    // Could be misread as "books rows missing from GSTR-2B" (layout 2) — must stay on
    // the categorized default instead of silently picking a direction.
    expect(detectGstReconLayout("what's missing in GSTR-2B")).toBe('categorized');
  });

  it('never swaps direction: books-then-portal and portal-then-books phrasings never collide', () => {
    const booksFlat = detectGstReconLayout('rows in my books not in the portal');
    const portalFlat = detectGstReconLayout('rows in the portal not in my books');
    expect(booksFlat).toBe('books_flat');
    expect(portalFlat).toBe('portal_flat');
    expect(booksFlat).not.toBe(portalFlat);
  });
});
