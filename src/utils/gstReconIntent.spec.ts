import { describe, expect, it } from 'vitest';
import {
  detectGstReconIntent,
  detectGstReconLayout,
  detectGstReconRelayoutTarget,
  extractRelayoutTargetSheetName,
  isCasualMissedRowsRecon,
  isGstReconPrompt,
  isGstReconRelayoutPrompt,
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

describe('isGstReconRelayoutPrompt — a FOLLOW-UP to change an already-created sheet, not a fresh run', () => {
  it('matches a change-layout verb plus layout wording', () => {
    expect(isGstReconRelayoutPrompt('change the layout of Missed vs GSTR-1 to a single table')).toBe(
      true,
    );
    expect(isGstReconRelayoutPrompt('switch Missed vs GSTR-2B to one table')).toBe(true);
    expect(isGstReconRelayoutPrompt('convert it to a flat list')).toBe(true);
    expect(isGstReconRelayoutPrompt('rearrange the sections into a single table')).toBe(true);
  });

  it('does NOT match an ordinary first-time layout request — that one has no change-layout verb', () => {
    expect(isGstReconRelayoutPrompt('Reconcile purchase register, just show the missed rows')).toBe(
      false,
    );
    expect(isGstReconRelayoutPrompt('reconcile sales register with a simple layout')).toBe(false);
  });

  it('does NOT match a change-layout verb with no layout wording, or layout wording with no verb', () => {
    expect(isGstReconRelayoutPrompt('change the GSTIN for this client')).toBe(false);
    expect(isGstReconRelayoutPrompt('the layout looks great')).toBe(false);
  });

  // Regression: these naturally-phrased follow-ups were previously rejected by this
  // gate before `detectGstReconRelayoutTarget` ever ran, because "make"/"put"/"group"
  // weren't recognized verbs and "by reason"/"portal flat"/"portal-only" weren't
  // recognized layout wording.
  it('matches "make"/"put ... back"/"group" as change-layout verbs', () => {
    expect(isGstReconRelayoutPrompt('make Missed vs GSTR-1 one table')).toBe(true);
    expect(isGstReconRelayoutPrompt('put the sections back')).toBe(true);
    expect(isGstReconRelayoutPrompt('group it by reason again')).toBe(true);
  });

  it('matches portal-direction wording ("portal-only" / "portal flat") as layout wording', () => {
    expect(isGstReconRelayoutPrompt('change it to portal-only rows')).toBe(true);
    expect(isGstReconRelayoutPrompt('switch to portal flat')).toBe(true);
  });

  // "make"/"put"/"group" are much more generic verbs than "change"/"switch"/"convert" —
  // confirms they still require the equally-specific layout wording to trigger, so an
  // unrelated sentence using one of them never false-positives into a relayout prompt.
  it('generic verbs ("make"/"put"/"group") never false-positive without layout wording', () => {
    expect(isGstReconRelayoutPrompt('make a new sheet called Summary')).toBe(false);
    expect(isGstReconRelayoutPrompt('put the GSTIN in column B')).toBe(false);
    expect(isGstReconRelayoutPrompt('group the vendors by region')).toBe(false);
  });
});

describe('detectGstReconRelayoutTarget', () => {
  it('defaults to books_flat for an unqualified "single table" / "one table" / "flat" ask', () => {
    expect(detectGstReconRelayoutTarget('change the layout to a single table')).toBe('books_flat');
    expect(detectGstReconRelayoutTarget('switch to one table')).toBe('books_flat');
  });

  it('selects categorized for "grouped" / "sections" / "by reason" / "categorized"', () => {
    expect(detectGstReconRelayoutTarget('switch back to categorized')).toBe('categorized');
    expect(detectGstReconRelayoutTarget('group it by reason again')).toBe('categorized');
    expect(detectGstReconRelayoutTarget('put the sections back')).toBe('categorized');
  });

  it('selects portal_flat for explicit portal-only wording', () => {
    expect(detectGstReconRelayoutTarget('change it to portal-only rows')).toBe('portal_flat');
    expect(detectGstReconRelayoutTarget('switch to portal flat')).toBe('portal_flat');
  });

  // Regression: "sections" can appear while describing what to GET RID OF, not what's
  // being asked for — "rearrange the SECTIONS into a single table" was previously
  // misread as 'categorized' because the bare "sections" mention won. An explicit
  // "single/one/flat table" phrase must take priority over a mere "sections" mention.
  it('an explicit "single/one/flat table" phrase wins over an incidental "sections" mention', () => {
    expect(detectGstReconRelayoutTarget('rearrange the sections into a single table')).toBe(
      'books_flat',
    );
  });
});

describe('extractRelayoutTargetSheetName', () => {
  it('extracts a "Missed vs X" sheet name from a relayout prompt', () => {
    expect(extractRelayoutTargetSheetName('change the layout of Missed vs GSTR-1 to a single table')).toBe(
      'Missed vs GSTR-1',
    );
    expect(extractRelayoutTargetSheetName('switch Missed vs GSTR-2B / GSTR-2A to one table.')).toBe(
      'Missed vs GSTR-2B / GSTR-2A',
    );
  });

  it('returns null when no sheet name is named', () => {
    expect(extractRelayoutTargetSheetName('change the layout to a single table')).toBeNull();
  });

  // Regression: "make Missed vs GSTR-1 one table" has no connector word ("to"/"into"/
  // "as") between the sheet name and the layout phrase — without a dedicated stop
  // condition for that, the match swallowed "one table" into the sheet name itself.
  it('stops before layout wording that directly abuts the sheet name with no connector word', () => {
    expect(extractRelayoutTargetSheetName('make Missed vs GSTR-1 one table')).toBe('Missed vs GSTR-1');
    expect(extractRelayoutTargetSheetName('reformat Missed vs GSTR-2B as a flat list')).toBe(
      'Missed vs GSTR-2B',
    );
  });
});
