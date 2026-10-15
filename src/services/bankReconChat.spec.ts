import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { RawTable } from '@/types/rawTable';

const LEDGER_HEADERS = ['Date', 'Particulars', 'Vch Type', 'Vch No.', 'Debit', 'Credit'];
const IMPORTED_HEADERS = ['Date', 'Value Date', 'Description', 'Ref No', 'Debit', 'Credit', 'Balance', 'Source', 'Flag'];
const FILE = 'bank_statement_sep2026_A_drcr_utf8bom.csv';

const rawTable: RawTable = {
  source: 'csv',
  layout: 'grid',
  fileName: FILE,
  rows: [{ ref: 'row 1', cells: [{ t: 'Date' }, { t: 'Narration' }] }],
};

const ledgerTop = { name: 'Bank Ledger', rows: [LEDGER_HEADERS, ['02-04-2025', 'ABC', 'Payment', 1, '', 100]] };
const importedTop = (name: string) => ({
  name,
  rows: [IMPORTED_HEADERS, ['2026-09-02', '2026-09-02', 'NEFT-ABC', 'X1', 100, '', 900, 'csv', '']],
});
const gridOf = (name: string) => ({ name, headers: [], values: [LEDGER_HEADERS, ['x']], rowCount: 2, columnCount: 6 });

const readSheetTops = vi.fn();
const readSheetsFull = vi.fn();
const sheetExists = vi.fn();
const runBankReconcile = vi.fn();

vi.mock('@/services/bankSheetDiscovery', async () => {
  const actual = await vi.importActual<typeof import('@/services/bankSheetDiscovery')>('@/services/bankSheetDiscovery');
  return { ...actual, readSheetTops: (...a: unknown[]) => readSheetTops(...a) };
});
vi.mock('@/services/gstSheetReader', () => ({
  readAllSheetHeaders: vi.fn().mockResolvedValue([]),
  readSheetsFull: (...a: unknown[]) => readSheetsFull(...a),
  sheetExists: (...a: unknown[]) => sheetExists(...a),
}));
vi.mock('@/services/bankReconService', async () => {
  const actual = await vi.importActual<typeof import('@/services/bankReconService')>('@/services/bankReconService');
  return { ...actual, runBankReconcile: (...a: unknown[]) => runBankReconcile(...a) };
});

const row = (over: Record<string, unknown>) => ({
  side: 'BOOKS',
  sheet_name: 'Bank Ledger',
  row: 8,
  date: '2026-09-02',
  direction: 'IN',
  amount: 38000,
  narration: 'Vasanth Super Mart',
  reference: '',
  reason: 'date_mismatch_outside_window',
  reason_label: 'Date differs from bank beyond clearing window',
  explanation: '',
  ...over,
});

const response = (over: Record<string, unknown> = {}) => ({
  job_id: 'j',
  status: 'complete',
  summary: { books_rows: 57, bank_rows: 65, matched: 43, matched_group: 2, matched_books_rows: 48, reversal_pairs: 0, books_only: 1, bank_only: 1, low_confidence_rows: 0, books_reasons: {}, bank_reasons: {} },
  rows: [
    row({}),
    row({ side: 'BANK', sheet_name: FILE, row: 18, external: true, source_ref: 'row 18', amount: 250, direction: 'OUT', narration: 'CHQ BOOK ISSUE CHG', reason_label: 'Bank charges not booked' }),
  ],
  actions: [{ type: 'CREATE_SHEET', sheetName: 'Bank Recon Summary', name: 'Bank Recon Summary' }],
  brs: { periodEnd: '2026-09-30', booksBalance: 1, bankBalance: 1, difference: 0, tiedOut: true, notes: [] },
  output_sheet_name: 'Bank Recon Summary',
  layout: 'categorized',
  layout_cache_id: 'c1',
  period_applied: { label: '2026-09-01 to 2026-09-30', start: '2026-09-01', end: '2026-09-30' },
  bank_source: { kind: 'attachment', file_name: FILE, file_type: 'csv', transactions: 65 },
  validation: { errors: [], warnings: [], notes: [] },
  ...over,
});

describe('bankReconChat — attached statement', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubGlobal('Excel', {});
    readSheetTops.mockReset().mockResolvedValue([ledgerTop]);
    readSheetsFull.mockReset().mockImplementation(async (names: string[]) => names.map(gridOf));
    sheetExists.mockReset().mockResolvedValue(false);
    runBankReconcile.mockReset().mockResolvedValue(response());
  });
  afterEach(() => vi.unstubAllGlobals());

  const attach = async (message = 'reconcile this') => {
    const { tryHandleBankReconWithAttachment } = await import('./bankReconChat');
    return tryHandleBankReconWithAttachment(message, rawTable);
  };

  it('first asks whether to create a sheet from the file — nothing is reconciled yet', async () => {
    const outcome = await attach();
    expect(outcome).toMatchObject({ kind: 'question' });
    if (outcome.kind !== 'question') throw new Error('unreachable');
    expect(outcome.options).toHaveLength(2);
    expect(outcome.options[0]).toMatch(/^Add as sheet “bank_statement_sep2026_A_drcr_u/);
    expect(outcome.options[1]).toBe('Use the file directly');
    // The question text must not itself contain either option, or an answer could be misread.
    expect(outcome.question).not.toMatch(/add as sheet|use the file directly/i);
    expect(runBankReconcile).not.toHaveBeenCalled();
  });

  it('"Use the file directly" reconciles the attached file; rows in the file become plain, non-clickable labels', async () => {
    await attach('reconcile bank ledger against this file');
    const { tryHandleBankReconChat } = await import('./bankReconChat');
    const outcome = await tryHandleBankReconChat('Use the file directly');

    expect(runBankReconcile).toHaveBeenCalledTimes(1);
    const req = runBankReconcile.mock.calls[0][0];
    expect(req.bank_statement.raw_table).toBe(rawTable);
    expect(req.bank_statement.data).toBeUndefined();
    expect(req.books_ledger.sheet_name).toBe('Bank Ledger');

    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') throw new Error('unreachable');
    const [books, bank] = outcome.missedRows!;
    // ledger row: still a clickable pill into the workbook
    expect(books).toMatchObject({ sheetName: 'Bank Ledger', row: 8, pillLabel: 'Books · row 8' });
    expect(books.inFile).toBeUndefined();
    // file row: no cell to jump to, and labelled as being in the file
    expect(bank).toMatchObject({ sheetName: null, row: null, inFile: true, pillLabel: 'row 18 in attached CSV' });
    expect(bank.tooltip).toMatch(/not in your workbook/);
  });

  it('"Add as sheet …" hands off to the statement import, named after the file, then reconciles against that sheet', async () => {
    await attach('reconcile this');
    const { tryHandleBankReconChat } = await import('./bankReconChat');
    // The app sends the answer wrapped like this (see answerQuestion in useConversation).
    const outcome = await tryHandleBankReconChat(
      'Replying to your question "How should I read the attached statement?" about my earlier request "Reconcile bank ledger against this file.": Add as sheet “bank_statement_sep2026_A_drcr_u”',
    );

    expect(outcome).toMatchObject({ kind: 'import_statement', preferredSheetName: expect.stringMatching(/^bank_statement_sep2026_A_drcr_u/) });
    if (outcome?.kind !== 'import_statement') throw new Error('unreachable');
    expect(outcome.preferredSheetName.length).toBeLessThanOrEqual(31);
    expect(outcome.rawTable).toBe(rawTable);
    // `{sheet}` is replaced with the real created name by the import step; the follow-up then names it explicitly.
    expect(outcome.followUpPrompt).toBe('Reconcile bank ledger against @[{sheet}]');
    expect(runBankReconcile).not.toHaveBeenCalled();
  });

  it('after the sheet exists, the follow-up reconciles against exactly that sheet (even next to another bank-looking sheet)', async () => {
    readSheetTops.mockResolvedValue([ledgerTop, importedTop('Bank Statement'), importedTop('bank_statement_sep2026_A_drcr_u')]);
    runBankReconcile.mockResolvedValue(response({ bank_source: { kind: 'sheet', file_name: 'x' }, rows: [row({}), row({ side: 'BANK', sheet_name: 'bank_statement_sep2026_A_drcr_u', row: 19 })] }));
    const { tryHandleBankReconChat } = await import('./bankReconChat');

    const outcome = await tryHandleBankReconChat('Reconcile bank ledger against @[bank_statement_sep2026_A_drcr_u]');

    const req = runBankReconcile.mock.calls[0][0];
    expect(req.bank_statement.sheet_name).toBe('bank_statement_sep2026_A_drcr_u');
    expect(req.bank_statement.raw_table).toBeUndefined();
    expect(outcome?.kind).toBe('recon_ready');
    if (outcome?.kind !== 'recon_ready') throw new Error('unreachable');
    // pills point into the created sheet and select the whole row
    expect(outcome.missedRows![1]).toMatchObject({ sheetName: 'bank_statement_sep2026_A_drcr_u', row: 19, selectFullRow: true });
  });

  it('without a mention, two bank-looking sheets are still ambiguous (the mention is what disambiguates)', async () => {
    readSheetTops.mockResolvedValue([ledgerTop, importedTop('Bank Statement'), importedTop('Bank Statement 2')]);
    const { tryHandleBankReconChat } = await import('./bankReconChat');
    const outcome = await tryHandleBankReconChat('Reconcile bank ledger against my statement');
    expect(outcome?.kind).toBe('question');
    expect(runBankReconcile).not.toHaveBeenCalled();
  });

  it('an unrelated reply to the question is not swallowed', async () => {
    await attach();
    const { tryHandleBankReconChat } = await import('./bankReconChat');
    expect(await tryHandleBankReconChat('what is the weather')).toBeNull();
  });

  it('offers the sheet a file was already added as, instead of creating a duplicate copy', async () => {
    const { tryHandleBankReconWithAttachment, tryHandleBankReconChat } = await import('./bankReconChat');

    // 1st time: add as a sheet. The import creates it (numbered if needed); the follow-up names it.
    await tryHandleBankReconWithAttachment('reconcile this', rawTable);
    const imp = await tryHandleBankReconChat('Add as sheet “bank_statement_sep2026_A_drcr_u”');
    expect(imp?.kind).toBe('import_statement');
    readSheetTops.mockResolvedValue([ledgerTop, importedTop('bank_statement_sep2026_A_drcr_u')]);
    await tryHandleBankReconChat('Reconcile bank ledger against @[bank_statement_sep2026_A_drcr_u]');
    // A sheet written from the file is held to the file's standard: the server refuses it if it was read incompletely.
    expect(runBankReconcile.mock.calls[0][0].bank_statement.strict).toBe(true);

    // 2nd time, same file: the existing sheet is offered first.
    const again = await tryHandleBankReconWithAttachment('reconcile this', rawTable);
    expect(again.kind).toBe('question');
    if (again.kind !== 'question') throw new Error('unreachable');
    expect(again.options).toHaveLength(3);
    expect(again.options[0]).toBe('Use existing sheet “bank_statement_sep2026_A_drcr_u”');
    expect(again.options[1]).toMatch(/^Add as sheet/);
    expect(again.options[2]).toBe('Use the file directly');
    expect(again.answer).toMatch(/already added/);

    // Choosing it reconciles against that sheet — no import, no raw file sent.
    runBankReconcile.mockClear();
    runBankReconcile.mockResolvedValue(response({ bank_source: { kind: 'sheet', file_name: 'x' }, rows: [row({})] }));
    const done = await tryHandleBankReconChat('Use existing sheet “bank_statement_sep2026_A_drcr_u”');
    expect(done?.kind).toBe('recon_ready');
    const req = runBankReconcile.mock.calls[0][0];
    expect(req.bank_statement.sheet_name).toBe('bank_statement_sep2026_A_drcr_u');
    expect(req.bank_statement.raw_table).toBeUndefined();
  });

  it('does not offer the earlier sheet once it has been deleted from the workbook', async () => {
    const { tryHandleBankReconWithAttachment, tryHandleBankReconChat } = await import('./bankReconChat');
    await tryHandleBankReconWithAttachment('reconcile this', rawTable);
    await tryHandleBankReconChat('Add as sheet “bank_statement_sep2026_A_drcr_u”');
    readSheetTops.mockResolvedValue([ledgerTop, importedTop('bank_statement_sep2026_A_drcr_u')]);
    await tryHandleBankReconChat('Reconcile bank ledger against @[bank_statement_sep2026_A_drcr_u]');

    readSheetTops.mockResolvedValue([ledgerTop]); // the user deleted it
    const again = await tryHandleBankReconWithAttachment('reconcile this', rawTable);
    if (again.kind !== 'question') throw new Error('expected a question');
    expect(again.options).toHaveLength(2);
  });

  it('a different file is never matched to an earlier sheet', async () => {
    const { tryHandleBankReconWithAttachment, tryHandleBankReconChat } = await import('./bankReconChat');
    await tryHandleBankReconWithAttachment('reconcile this', rawTable);
    await tryHandleBankReconChat('Add as sheet “bank_statement_sep2026_A_drcr_u”');
    readSheetTops.mockResolvedValue([ledgerTop, importedTop('bank_statement_sep2026_A_drcr_u')]);
    await tryHandleBankReconChat('Reconcile bank ledger against @[bank_statement_sep2026_A_drcr_u]');

    const other: RawTable = { ...rawTable, fileName: 'october.csv', rows: [{ ref: 'row 1', cells: [{ t: 'Date' }, { t: 'x' }] }, ...rawTable.rows] };
    const q = await tryHandleBankReconWithAttachment('reconcile this', other);
    if (q.kind !== 'question') throw new Error('expected a question');
    expect(q.options).toHaveLength(2);
  });

  it('when the summary sheet name is taken, the collision card still carries the clickable unreconciled list', async () => {
    sheetExists.mockResolvedValue(true);
    const { tryHandleBankReconWithAttachment, tryHandleBankReconChat } = await import('./bankReconChat');
    await tryHandleBankReconWithAttachment('reconcile this', rawTable);
    const outcome = await tryHandleBankReconChat('Use the file directly');
    expect(outcome?.kind).toBe('sheet_collision');
    if (outcome?.kind !== 'sheet_collision') throw new Error('unreachable');
    expect(outcome.sheetName).toBe('Bank Recon Summary');
    expect(outcome.missedRows).toHaveLength(2);
    expect(outcome.missedRows![0]).toMatchObject({ sheetName: 'Bank Ledger', row: 8 });
    expect(outcome.missedRows![1]).toMatchObject({ inFile: true, sheetName: null });
    expect(outcome.missedRowsTitle).toBe('Unreconciled items');
  });

  it('a PDF reads exactly like a CSV: same question, and its in-file rows are labelled by page and line', async () => {
    const pdf: RawTable = { source: 'pdf', layout: 'positioned', fileName: 'bank_statement_sep2026_digital.pdf', rows: rawTable.rows };
    runBankReconcile.mockResolvedValue(
      response({
        bank_source: { kind: 'attachment', file_name: pdf.fileName, file_type: 'pdf', transactions: 65 },
        rows: [row({}), row({ side: 'BANK', sheet_name: pdf.fileName, row: 14, external: true, source_ref: 'p2 l14', amount: 250, direction: 'OUT' })],
      }),
    );
    const { tryHandleBankReconWithAttachment, tryHandleBankReconChat } = await import('./bankReconChat');

    const q = await tryHandleBankReconWithAttachment('reconcile bank ledger against this file', pdf);
    expect(q.kind).toBe('question');
    if (q.kind !== 'question') throw new Error('unreachable');
    expect(q.options[q.options.length - 1]).toBe('Use the file directly');
    expect(q.options[0]).toMatch(/^Add as sheet “bank_statement_sep2026_digital/);

    const out = await tryHandleBankReconChat('Use the file directly');
    if (out?.kind !== 'recon_ready') throw new Error('expected recon_ready');
    expect(out.missedRows![1]).toMatchObject({ inFile: true, sheetName: null, pillLabel: 'p2 l14 in attached PDF' });
    expect(runBankReconcile.mock.calls[0][0].bank_statement.raw_table).toBe(pdf);
  });
});
