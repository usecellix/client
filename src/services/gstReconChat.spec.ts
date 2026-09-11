import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { SheetAction } from '@/types/sheet-actions';
import type { GstReconcileResponse } from '@/services/gstReconService';

const PR_SHEET = 'Purchase Register';
const PORTAL_SHEET = 'GSTR-2B';
const OUTPUT_SHEET = 'Missed vs GSTR-2B';

const PR_HEADERS = [
  'Supplier GSTIN',
  'Invoice No',
  'Invoice Date',
  'Taxable Value',
  'CGST',
  'SGST',
  'IGST',
  'Narration',
];

const PORTAL_HEADERS = [
  'GSTIN of Supplier',
  'Invoice Number',
  'Invoice Date',
  'Taxable Value',
  'Document Type',
  'ITC Available',
  'CGST',
  'SGST',
  'IGST',
];

function grid(name: string, headers: string[]) {
  return {
    name,
    headers,
    values: [headers, headers.map((_, i) => `r1c${i}`), headers.map((_, i) => `r2c${i}`)],
    rowCount: 3,
    columnCount: headers.length,
  };
}

function fixtureActions(sheetName: string): SheetAction[] {
  return [
    {
      type: 'CREATE_SHEET',
      sheetName,
      name: sheetName,
      relativeTo: PR_SHEET,
      position: 'after',
    },
    {
      type: 'WRITE_TABLE',
      sheetName,
      headers: ['GSTIN', 'Vendor Name'],
      rows: [['29ABCDE1234F1Z5', 'Acme Co']],
    },
  ];
}

function buildMockResult(overrides: Partial<GstReconcileResponse> = {}): GstReconcileResponse {
  return {
    job_id: 'job-1',
    status: 'complete',
    reconciliation_type: 'PR_VS_GSTR2B',
    client_gstin: null,
    summary: {
      total_pr_rows: 10,
      total_portal_rows: 9,
      total_ims_rows: 0,
      exact_matched: 8,
      partial_matched: 0,
      credit_notes: 0,
      pr_only: 2,
      portal_only: 0,
      ims_rejected: 0,
      ims_pending: 0,
      ims_auto_accept: 0,
      rcm_flagged: 0,
      itc_matched: 0,
      itc_at_risk: 0,
      rcm_payable: 0,
      itc_ims_rejected: 0,
      itc_ims_pending: 0,
      matched_exact: 8,
      matched_fallback: 0,
      mismatch_blank_gstin: 0,
      mismatch_blank_gstin_likely_matched: 0,
      mismatch_blank_taxable_value: 0,
      mismatch_ambiguous_rate_slab: 0,
      mismatch_amount: 0,
      mismatch_date: 0,
      mismatch_genuinely_missing: 2,
      gstin_mismatch_count: 0,
    },
    rows: [],
    actions: fixtureActions(OUTPUT_SHEET),
    confidence: 1,
    exceptions: [],
    output_sheet_name: OUTPUT_SHEET,
    audit_log_id: 'audit-1',
    missed_books_only: true,
    ...overrides,
  };
}

const readAllSheetHeaders = vi.fn();
const readSheetsFull = vi.fn();
const runGstReconcile = vi.fn();

vi.mock('@/services/gstSheetReader', () => ({
  readAllSheetHeaders: (...args: unknown[]) => readAllSheetHeaders(...args),
  readSheetsFull: (...args: unknown[]) => readSheetsFull(...args),
}));

vi.mock('@/services/gstReconService', async () => {
  const actual = await vi.importActual<typeof import('@/services/gstReconService')>(
    '@/services/gstReconService',
  );
  return {
    ...actual,
    runGstReconcile: (...args: unknown[]) => runGstReconcile(...args),
  };
});

describe('gstReconChat — output sheet collision', () => {
  beforeEach(() => {
    // gstReconChat.ts keeps its pending-collision/pending-context state in module-level
    // `let`s — reset the module registry so each test starts from a clean singleton
    // instead of leaking state (e.g. an unresolved collision) into the next test.
    vi.resetModules();
    vi.stubGlobal('Excel', {});
    readAllSheetHeaders.mockReset();
    readSheetsFull.mockReset();
    runGstReconcile.mockReset();

    readSheetsFull.mockResolvedValue([grid(PR_SHEET, PR_HEADERS), grid(PORTAL_SHEET, PORTAL_HEADERS)]);
    runGstReconcile.mockResolvedValue(buildMockResult());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  async function runCasualRecon() {
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    return tryHandleGstReconChat('Reconcile purchase register');
  }

  async function triggerCollision() {
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
      { name: OUTPUT_SHEET, headers: ['GSTIN', 'Vendor Name'] },
    ]);
    const outcome = await runCasualRecon();
    if (outcome?.kind !== 'sheet_collision') {
      throw new Error(`expected sheet_collision, got ${outcome?.kind}`);
    }
    return outcome;
  }

  it('collision emits a button-choice payload — not a plain-text question expecting a typed answer', async () => {
    const outcome = await triggerCollision();

    expect(outcome.kind).toBe('sheet_collision');
    expect(outcome.sheetName).toBe(OUTPUT_SHEET);
    expect(typeof outcome.collisionId).toBe('string');
    expect(outcome.collisionId.length).toBeGreaterThan(0);
    // The answer carries the reconciliation breakdown for context, not a
    // "type Overwrite or Create new" instruction — there is nothing to parse.
    expect(outcome.answer).not.toMatch(/type\s+"?overwrite/i);
  });

  it('a message that looks like the old typed reply ("Overwrite") is no longer treated as a reply', async () => {
    await triggerCollision();

    // Previously this text-matching hack resolved the pending collision. Now that
    // resolution only happens through resolveGstReconSheetCollision (button click),
    // a plain chat message must fall back to ordinary intent handling instead.
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat('Overwrite');

    expect(outcome).toBeNull();
  });

  it('clicking "Overwrite" deletes and recreates the existing sheet — no chat text involved', async () => {
    const collision = await triggerCollision();

    const { resolveGstReconSheetCollision } = await import('./gstReconChat');
    const outcome = await resolveGstReconSheetCollision(collision.collisionId, 'overwrite');

    expect(outcome.kind).toBe('recon_ready');
    if (outcome.kind !== 'recon_ready') return;

    // DELETE_SHEET first, so CREATE_SHEET/WRITE_TABLE land on a fresh, empty
    // sheet — no in-place row-0 clear that would trip the header guard, and
    // no OverwriteGuard trip on WRITE_TABLE since nothing pre-exists there.
    expect(outcome.actions[0]).toEqual({ type: 'DELETE_SHEET', sheetName: OUTPUT_SHEET });
    // Original CREATE_SHEET/WRITE_TABLE actions follow, unchanged, same sheet name.
    expect(outcome.actions.slice(1)).toEqual(fixtureActions(OUTPUT_SHEET));
    // sheetName lets the caller confirm "Sheet created as X" without re-deriving
    // the name from the action array.
    expect(outcome.sheetName).toBe(OUTPUT_SHEET);
  });

  it('"Overwrite" never emits a CLEAR_RANGE/FORMAT_RANGE/CLEAR_ALL touching the old sheet\'s row 0', async () => {
    const collision = await triggerCollision();

    const { resolveGstReconSheetCollision } = await import('./gstReconChat');
    const outcome = await resolveGstReconSheetCollision(collision.collisionId, 'overwrite');

    expect(outcome.kind).toBe('recon_ready');
    if (outcome.kind !== 'recon_ready') return;

    const touchesOldRowZero = outcome.actions.some((action) => {
      if (!['CLEAR_RANGE', 'CLEAR_CONTENT', 'CLEAR_FORMAT', 'CLEAR_ALL', 'FORMAT_RANGE'].includes(action.type)) {
        return false;
      }
      if (action.sheetName !== OUTPUT_SHEET) return false;
      const rowStart = action.row ?? 0;
      const rangeStartsAtRow1 = typeof action.range === 'string' && /^[A-Z]+1(:|$)/.test(action.range);
      return rowStart === 0 || rangeStartsAtRow1;
    });
    expect(touchesOldRowZero).toBe(false);
  });

  it('clicking "Create new" writes into a versioned sheet name instead — no chat text involved', async () => {
    const collision = await triggerCollision();
    // Re-checked inside the resolver — still only the original name is taken.
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
      { name: OUTPUT_SHEET, headers: ['GSTIN', 'Vendor Name'] },
    ]);

    const { resolveGstReconSheetCollision } = await import('./gstReconChat');
    const outcome = await resolveGstReconSheetCollision(collision.collisionId, 'new');

    expect(outcome.kind).toBe('recon_ready');
    if (outcome.kind !== 'recon_ready') return;

    const versionedName = `${OUTPUT_SHEET} (2)`;
    expect(outcome.actions).toEqual(fixtureActions(versionedName));
    expect(outcome.answer).toContain(versionedName);
    expect(outcome.sheetName).toBe(versionedName);
  });

  it('a stale/unknown collisionId is refused rather than silently resolved', async () => {
    await triggerCollision();

    const { resolveGstReconSheetCollision } = await import('./gstReconChat');
    const outcome = await resolveGstReconSheetCollision('collision_does_not_exist', 'overwrite');

    expect(outcome.kind).toBe('message_only');
  });

  it('resolving once clears the pending collision — a second click on the same id is refused', async () => {
    const collision = await triggerCollision();

    const { resolveGstReconSheetCollision } = await import('./gstReconChat');
    const first = await resolveGstReconSheetCollision(collision.collisionId, 'overwrite');
    expect(first.kind).toBe('recon_ready');

    const second = await resolveGstReconSheetCollision(collision.collisionId, 'overwrite');
    expect(second.kind).toBe('message_only');
  });

  it('no collision: behavior is unchanged, goes straight to the summary', async () => {
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
    ]);

    const outcome = await runCasualRecon();

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.actions).toEqual(fixtureActions(OUTPUT_SHEET));
    expect(outcome.sheetName).toBe(OUTPUT_SHEET);
  });
});
