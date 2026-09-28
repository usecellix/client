import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { SheetAction } from '@/types/sheet-actions';
import type { GstReconcileResponse } from '@/services/gstReconService';

const PR_SHEET = 'Purchase Register';
const PORTAL_SHEET = 'GSTR-2B';
const PORTAL_2A_SHEET = 'GSTR-2A';
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
const sheetExists = vi.fn();
const runGstReconcile = vi.fn();

vi.mock('@/services/gstSheetReader', () => ({
  readAllSheetHeaders: (...args: unknown[]) => readAllSheetHeaders(...args),
  readSheetsFull: (...args: unknown[]) => readSheetsFull(...args),
  sheetExists: (...args: unknown[]) => sheetExists(...args),
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
    sheetExists.mockReset();
    runGstReconcile.mockReset();

    readSheetsFull.mockResolvedValue([grid(PR_SHEET, PR_HEADERS), grid(PORTAL_SHEET, PORTAL_HEADERS)]);
    // Default: nothing collides. Individual tests override this to simulate
    // headerIndex missing a sheet that genuinely exists.
    sheetExists.mockResolvedValue(false);
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

  it('resolving a sheet-name collision still carries the missedRows list on the recon_ready outcome — regression for the missed-rows card disappearing after Overwrite/Create-new', async () => {
    runGstReconcile.mockResolvedValue(
      buildMockResult({
        rows: [
          {
            status: 'PR_ONLY',
            invoice_number: '',
            gstin: '',
            vendor_name: 'Coral Bay Logistics',
            mismatch_reason: 'blank_counterparty_gstin',
            explanation: 'GSTIN is blank in the register for this row — cannot be matched.',
            books_sheet_name: PR_SHEET,
            books_row: 6,
          },
        ],
      }),
    );
    const collision = await triggerCollision();

    const { resolveGstReconSheetCollision } = await import('./gstReconChat');
    const outcome = await resolveGstReconSheetCollision(collision.collisionId, 'overwrite');

    expect(outcome.kind).toBe('recon_ready');
    if (outcome.kind !== 'recon_ready') return;
    expect(outcome.missedRows).toHaveLength(1);
    expect(outcome.missedRows?.[0]).toMatchObject({
      vendorName: 'Coral Bay Logistics',
      gstin: '',
      reason: 'blank_counterparty_gstin',
      sheetName: PR_SHEET,
      row: 6,
    });
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

  it('collision detection still catches an existing sheet name when headerIndex misses it (e.g. a stale/partial context read) — the fresh sheetExists check is authoritative', async () => {
    // headerIndex (a snapshot read earlier in the function) does NOT include the
    // output sheet — simulating a partial/stale read that missed a sheet other
    // than the active one. sheetExists, a separate fresh/minimal live check,
    // still says it's there — collision must still be caught.
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
    ]);
    sheetExists.mockImplementation((name: string) => Promise.resolve(name === OUTPUT_SHEET));

    const outcome = await runCasualRecon();

    expect(outcome?.kind).toBe('sheet_collision');
    if (outcome?.kind !== 'sheet_collision') return;
    expect(outcome.sheetName).toBe(OUTPUT_SHEET);
    expect(sheetExists).toHaveBeenCalledWith(OUTPUT_SHEET);
  });

  it('sheetExists is skipped when headerIndex already found the collision (no redundant live check)', async () => {
    await triggerCollision();
    // triggerCollision's headerIndex already includes OUTPUT_SHEET, so the
    // short-circuiting `||` must never even call the live check.
    expect(sheetExists).not.toHaveBeenCalled();
  });
});

describe('gstReconChat — missedRows on the recon_ready outcome (clickable jump-to-cell list)', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubGlobal('Excel', {});
    readAllSheetHeaders.mockReset();
    readSheetsFull.mockReset();
    sheetExists.mockReset();
    runGstReconcile.mockReset();

    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
    ]);
    readSheetsFull.mockResolvedValue([grid(PR_SHEET, PR_HEADERS), grid(PORTAL_SHEET, PORTAL_HEADERS)]);
    sheetExists.mockResolvedValue(false);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('carries one entry per PR_ONLY row, with the real books sheet name and row number for jumping to it', async () => {
    runGstReconcile.mockResolvedValue(
      buildMockResult({
        rows: [
          {
            status: 'MATCHED',
            invoice_number: 'INV-1',
            gstin: '29ABCDE1234F1Z5',
          },
          {
            status: 'PR_ONLY',
            invoice_number: 'INV-2',
            gstin: '29XYZAB5678C1Z9',
            vendor_name: 'Acme Traders',
            mismatch_reason: 'date_mismatch',
            explanation: 'Same GSTIN and amount found in portal, but date differs.',
            books_sheet_name: PR_SHEET,
            books_row: 14,
          },
        ],
      }),
    );

    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat('Reconcile purchase register');

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    // Only the PR_ONLY row — the MATCHED row is never included.
    expect(outcome.missedRows).toHaveLength(1);
    expect(outcome.missedRows?.[0]).toEqual({
      vendorName: 'Acme Traders',
      gstin: '29XYZAB5678C1Z9',
      reason: 'date_mismatch',
      explanation: 'Same GSTIN and amount found in portal, but date differs.',
      sheetName: PR_SHEET,
      row: 14,
    });
  });

  it('falls back to the invoice number (or "Unnamed row") and empty GSTIN when the row is missing those fields — never a bare/undefined display value', async () => {
    runGstReconcile.mockResolvedValue(
      buildMockResult({
        rows: [
          {
            status: 'PR_ONLY',
            invoice_number: '',
            gstin: '',
            vendor_name: '',
            mismatch_reason: 'blank_counterparty_gstin',
            books_sheet_name: PR_SHEET,
            books_row: 6,
          },
        ],
      }),
    );

    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat('Reconcile purchase register');

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.missedRows?.[0].vendorName).toBe('Unnamed row');
    expect(outcome.missedRows?.[0].gstin).toBe('');
  });

  it('has a null sheetName/row when the row has no books-side reference at all (never crashes, never fabricates a cell)', async () => {
    runGstReconcile.mockResolvedValue(
      buildMockResult({
        rows: [
          {
            status: 'PR_ONLY',
            invoice_number: 'INV-3',
            gstin: '29ABCDE1234F1Z5',
            vendor_name: 'Some Vendor',
            mismatch_reason: 'genuinely_missing',
          },
        ],
      }),
    );

    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat('Reconcile purchase register');

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.missedRows?.[0].sheetName).toBeNull();
    expect(outcome.missedRows?.[0].row).toBeNull();
  });
});

describe('gstReconChat — write-time collision safety net (registerGstReconWriteFailureCollision)', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubGlobal('Excel', {});
    readAllSheetHeaders.mockReset();
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
      { name: OUTPUT_SHEET, headers: ['GSTIN', 'Vendor Name'] },
    ]);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('registers a write-time collision and resolves "overwrite" by deleting and recreating the same sheet — no reconciliation re-run needed', async () => {
    const { registerGstReconWriteFailureCollision, resolveGstReconSheetCollision } = await import(
      './gstReconChat'
    );
    const actions = fixtureActions(OUTPUT_SHEET);
    const { collisionId, sheetName } = registerGstReconWriteFailureCollision(actions, OUTPUT_SHEET);
    expect(sheetName).toBe(OUTPUT_SHEET);
    expect(collisionId).toMatch(/^write_collision_/);

    const outcome = await resolveGstReconSheetCollision(collisionId, 'overwrite');
    expect(outcome.kind).toBe('recon_ready');
    if (outcome.kind !== 'recon_ready') return;
    expect(outcome.actions[0]).toEqual({ type: 'DELETE_SHEET', sheetName: OUTPUT_SHEET });
    expect(outcome.actions.slice(1)).toEqual(actions);
    expect(outcome.sheetName).toBe(OUTPUT_SHEET);
  });

  it('resolves "new" by renaming to a versioned sheet name — checked fresh against the live workbook', async () => {
    const { registerGstReconWriteFailureCollision, resolveGstReconSheetCollision } = await import(
      './gstReconChat'
    );
    const actions = fixtureActions(OUTPUT_SHEET);
    const { collisionId } = registerGstReconWriteFailureCollision(actions, OUTPUT_SHEET);

    const outcome = await resolveGstReconSheetCollision(collisionId, 'new');
    expect(outcome.kind).toBe('recon_ready');
    if (outcome.kind !== 'recon_ready') return;
    const versionedName = `${OUTPUT_SHEET} (2)`;
    expect(outcome.actions).toEqual(fixtureActions(versionedName));
    expect(outcome.sheetName).toBe(versionedName);
  });

  it('a stale/unknown write-time collisionId is refused, never silently resolved', async () => {
    const { resolveGstReconSheetCollision } = await import('./gstReconChat');
    const outcome = await resolveGstReconSheetCollision('write_collision_does_not_exist', 'overwrite');
    expect(outcome.kind).toBe('message_only');
  });

  it('write-time and chat-detected collisions never collide with each other — resolving one never touches the other\'s pending state', async () => {
    const { registerGstReconWriteFailureCollision, resolveGstReconSheetCollision } = await import(
      './gstReconChat'
    );
    const actions = fixtureActions(OUTPUT_SHEET);
    const { collisionId: writeId } = registerGstReconWriteFailureCollision(actions, OUTPUT_SHEET);

    // Resolving a DIFFERENT, chat-detected-style id must not find (or consume) the
    // write-time entry, and must not crash — it should cleanly report "not available".
    const wrongKind = await resolveGstReconSheetCollision('collision_999', 'overwrite');
    expect(wrongKind.kind).toBe('message_only');

    // The real write-time collision is still resolvable afterward.
    const real = await resolveGstReconSheetCollision(writeId, 'overwrite');
    expect(real.kind).toBe('recon_ready');
  });
});

function sheetNameForLayout(layout?: string): string {
  if (layout === 'books_flat') return `${OUTPUT_SHEET} — Books Only`;
  if (layout === 'portal_flat') return `${OUTPUT_SHEET} — Portal Only`;
  return OUTPUT_SHEET;
}

describe('gstReconChat — collision detection is scoped per layout-specific sheet name', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubGlobal('Excel', {});
    readAllSheetHeaders.mockReset();
    readSheetsFull.mockReset();
    sheetExists.mockReset();
    runGstReconcile.mockReset();

    readSheetsFull.mockResolvedValue([grid(PR_SHEET, PR_HEADERS), grid(PORTAL_SHEET, PORTAL_HEADERS)]);
    sheetExists.mockResolvedValue(false);
    runGstReconcile.mockImplementation(
      (req: { layout?: string }) =>
        Promise.resolve(
          buildMockResult({
            output_sheet_name: sheetNameForLayout(req.layout),
            actions: fixtureActions(sheetNameForLayout(req.layout)),
            // portal_flat's own row set is portal-only rows — give it something to
            // report so this scenario reaches recon_ready instead of "no rows found".
            summary: {
              ...buildMockResult().summary,
              portal_only: req.layout === 'portal_flat' ? 3 : 0,
            },
          }),
        ) as never,
    );
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('a different layout\'s sheet already existing does NOT trigger the collision card', async () => {
    // Only the categorized sheet exists — books_flat's own name ("... — Books Only")
    // does not, so switching to books_flat must create a sibling sheet, not collide.
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
      { name: OUTPUT_SHEET, headers: ['GSTIN', 'Vendor Name'] },
    ]);

    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat(
      'Reconcile purchase register, just show me the missed rows',
    );

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.sheetName).toBe(`${OUTPUT_SHEET} — Books Only`);
  });

  it('the SAME layout\'s own sheet name already existing DOES trigger the collision card', async () => {
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
      { name: `${OUTPUT_SHEET} — Books Only`, headers: ['GSTIN', 'Vendor Name'] },
    ]);

    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat(
      'Reconcile purchase register, just show me the missed rows',
    );

    expect(outcome?.kind).toBe('sheet_collision');
    if (outcome?.kind !== 'sheet_collision') return;
    expect(outcome.sheetName).toBe(`${OUTPUT_SHEET} — Books Only`);
  });

  it('portal_flat and books_flat never collide with each other even when both already exist as sheets — each still resolves to its own run', async () => {
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
      { name: `${OUTPUT_SHEET} — Books Only`, headers: ['GSTIN', 'Vendor Name'] },
    ]);

    const { tryHandleGstReconChat } = await import('./gstReconChat');
    // portal_flat's own name ("... — Portal Only") is not among the existing sheets,
    // even though books_flat's name is — must not collide.
    const outcome = await tryHandleGstReconChat(
      'Reconcile purchase register — rows in the portal not in my books',
    );

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.sheetName).toBe(`${OUTPUT_SHEET} — Portal Only`);
  });
});

class MemoryStorage implements Storage {
  private store = new Map<string, string>();
  get length(): number {
    return this.store.size;
  }
  clear(): void {
    this.store.clear();
  }
  getItem(key: string): string | null {
    return this.store.get(key) ?? null;
  }
  key(index: number): string | null {
    return Array.from(this.store.keys())[index] ?? null;
  }
  removeItem(key: string): void {
    this.store.delete(key);
  }
  setItem(key: string, value: string): void {
    this.store.set(key, value);
  }
}

describe('gstReconChat — GSTR-2B/2A portal source preference (ask once, remember)', () => {
  const WORKBOOK_KEY = 'Client_Books.xlsx';

  beforeEach(() => {
    vi.resetModules();
    vi.stubGlobal('Excel', {});
    vi.stubGlobal('localStorage', new MemoryStorage());
    readAllSheetHeaders.mockReset();
    readSheetsFull.mockReset();
    sheetExists.mockReset();
    runGstReconcile.mockReset();

    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
      { name: PORTAL_2A_SHEET, headers: PORTAL_HEADERS },
    ]);
    readSheetsFull.mockImplementation((names: string[]) =>
      Promise.resolve(
        names.map((n) =>
          n === PR_SHEET ? grid(PR_SHEET, PR_HEADERS) : grid(n, PORTAL_HEADERS),
        ),
      ),
    );
    sheetExists.mockResolvedValue(false);
    runGstReconcile.mockImplementation(
      (req: { portal_file?: { sheet_name: string }; portal_file_2a?: { sheet_name: string } }) => {
        const usedBoth = Boolean(req.portal_file_2a);
        const sheetName = usedBoth
          ? 'Missed vs GSTR-2B GSTR-2A'
          : req.portal_file?.sheet_name === PORTAL_2A_SHEET
            ? 'Missed vs GSTR-2A'
            : 'Missed vs GSTR-2B';
        return Promise.resolve(
          buildMockResult({
            output_sheet_name: sheetName,
            actions: fixtureActions(sheetName),
          }),
        );
      },
    );
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('an explicit phrase ("2b only") resolves immediately with no ask, and persists the choice', async () => {
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat('Reconcile purchase register against 2b only', {
      workbookKey: WORKBOOK_KEY,
    });

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.sheetName).toBe('Missed vs GSTR-2B');
    expect(runGstReconcile).toHaveBeenCalledTimes(1);
    const call = runGstReconcile.mock.calls[0][0];
    expect(call.portal_file.sheet_name).toBe(PORTAL_SHEET);
    expect(call.portal_file_2a).toBeUndefined();

    const { loadChatSessions } = await import('@/utils/chatSessionStorage');
    expect(loadChatSessions(WORKBOOK_KEY)?.gstPurchasePortalPreference).toBe('gstr2b_only');
  });

  it('no explicit phrase and no stored preference produces a docked "question" outcome with the three plain-label options', async () => {
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat('Reconcile purchase register', {
      workbookKey: WORKBOOK_KEY,
    });

    expect(outcome?.kind).toBe('question');
    if (outcome?.kind !== 'question') return;
    expect(outcome.question).toContain('GSTR-2B');
    expect(outcome.question).toContain('GSTR-2A');
    expect(outcome.options).toEqual(['GSTR-2B only', 'GSTR-2A only', 'Both, combined']);
    expect(runGstReconcile).not.toHaveBeenCalled();
  });

  it('resolves via the exact wrapped reply useConversation.ts sends when a docked question is answered', async () => {
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const asked = await tryHandleGstReconChat('Reconcile purchase register', {
      workbookKey: WORKBOOK_KEY,
    });
    expect(asked?.kind).toBe('question');
    if (asked?.kind !== 'question') return;

    // This is the literal shape `answerQuestion` in useConversation.ts sends back — it
    // embeds the question text (which itself mentions both GSTR-2B and GSTR-2A) ahead of
    // the actual answer. This is the regression test for the bug: naively scanning the
    // whole payload for "2b"/"2a" mentions would see both and refuse to resolve.
    const wrapped = `Replying to your question "${asked.question}" about my earlier request "Reconcile purchase register": GSTR-2B only`;
    const outcome = await tryHandleGstReconChat(wrapped, { workbookKey: WORKBOOK_KEY });

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.sheetName).toBe('Missed vs GSTR-2B');
    const call = runGstReconcile.mock.calls[0][0];
    expect(call.portal_file.sheet_name).toBe(PORTAL_SHEET);
    expect(call.portal_file_2a).toBeUndefined();

    const { loadChatSessions } = await import('@/utils/chatSessionStorage');
    expect(loadChatSessions(WORKBOOK_KEY)?.gstPurchasePortalPreference).toBe('gstr2b_only');
  });

  it('a wrapped reply answering "GSTR-2A only" runs against GSTR-2A only', async () => {
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const asked = await tryHandleGstReconChat('Reconcile purchase register', {
      workbookKey: WORKBOOK_KEY,
    });
    if (asked?.kind !== 'question') throw new Error('expected question');

    const wrapped = `Replying to your question "${asked.question}" about my earlier request "Reconcile purchase register": GSTR-2A only`;
    const outcome = await tryHandleGstReconChat(wrapped, { workbookKey: WORKBOOK_KEY });

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.sheetName).toBe('Missed vs GSTR-2A');
    const call = runGstReconcile.mock.calls[0][0];
    expect(call.portal_file.sheet_name).toBe(PORTAL_2A_SHEET);
  });

  it('a wrapped reply answering "Both, combined" runs the dual-source union', async () => {
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const asked = await tryHandleGstReconChat('Reconcile purchase register', {
      workbookKey: WORKBOOK_KEY,
    });
    if (asked?.kind !== 'question') throw new Error('expected question');

    const wrapped = `Replying to your question "${asked.question}" about my earlier request "Reconcile purchase register": Both, combined`;
    const outcome = await tryHandleGstReconChat(wrapped, { workbookKey: WORKBOOK_KEY });

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.sheetName).toBe('Missed vs GSTR-2B GSTR-2A');
    const call = runGstReconcile.mock.calls[0][0];
    expect(call.portal_file_2a).toBeDefined();
  });

  it('resolves on the first answer even when workbookKey is unavailable (persistence is a side effect, never a dependency of resuming)', async () => {
    // Regression test: real usage showed the docked question being re-asked a second
    // time after the user had already answered it once. Root cause — the resume path
    // persisted the choice then RELOADED it from storage on the recursive call; if that
    // round-trip ever failed silently (e.g. workbookKey unavailable), the resumed run
    // found no preference and asked again. Passing no workbookKey here reproduces exactly
    // that failure mode for the persistence step, and the fix (passing the resolved
    // choice straight through instead of relying on storage) must still resolve first try.
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const asked = await tryHandleGstReconChat('Reconcile purchase register');
    expect(asked?.kind).toBe('question');
    if (asked?.kind !== 'question') return;

    const wrapped = `Replying to your question "${asked.question}" about my earlier request "Reconcile purchase register": GSTR-2B only`;
    const outcome = await tryHandleGstReconChat(wrapped);

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.sheetName).toBe('Missed vs GSTR-2B');
    expect(runGstReconcile).toHaveBeenCalledTimes(1);
  });

  it('an unrelated message while the ask is pending falls through to ordinary handling instead of getting stuck', async () => {
    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const asked = await tryHandleGstReconChat('Reconcile purchase register', {
      workbookKey: WORKBOOK_KEY,
    });
    expect(asked?.kind).toBe('question');

    const reply = await tryHandleGstReconChat('what is my total revenue this month', {
      workbookKey: WORKBOOK_KEY,
    });
    // Not GST-recon phrasing and not a resolvable answer to the pending ask — falls
    // through to null (the caller's generic chat path), matching the real dispatch gate.
    expect(reply).toBeNull();
    expect(runGstReconcile).not.toHaveBeenCalled();
  });

  it('a stored preference short-circuits the ask on a subsequent run — no re-ask', async () => {
    const { saveChatSessions } = await import('@/utils/chatSessionStorage');
    saveChatSessions(WORKBOOK_KEY, {
      activeSessionId: null,
      sessions: [],
      gstPurchasePortalPreference: 'gstr2a_only',
    });

    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat('Reconcile purchase register', {
      workbookKey: WORKBOOK_KEY,
    });

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') return;
    expect(outcome.sheetName).toBe('Missed vs GSTR-2A');
  });

  it('only one portal sheet present never triggers the ask, regardless of workbookKey', async () => {
    readAllSheetHeaders.mockResolvedValue([
      { name: PR_SHEET, headers: PR_HEADERS },
      { name: PORTAL_SHEET, headers: PORTAL_HEADERS },
    ]);
    readSheetsFull.mockResolvedValue([grid(PR_SHEET, PR_HEADERS), grid(PORTAL_SHEET, PORTAL_HEADERS)]);

    const { tryHandleGstReconChat } = await import('./gstReconChat');
    const outcome = await tryHandleGstReconChat('Reconcile purchase register', {
      workbookKey: WORKBOOK_KEY,
    });

    expect(outcome?.kind).toBe('recon_ready');
  });
});
