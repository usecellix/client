import { afterEach, describe, expect, it, vi } from 'vitest';
import { partitionActions } from '@/engine/actionNormalizer';
import { INTERNAL_COPY_MARKERS } from '@/utils/userFacingResponse';
import { shortenActionPreviewCopy } from '@/utils/actionPreviewCopy';
import { SheetAction } from '@/types/sheet-actions';
import {
  BankStatementImportError,
  BankStatementImportResult,
  buildBankStatementAnswer,
  buildBankStatementSummary,
  describeBankStatementImportFailure,
  describeVerification,
  findFreeSheetName,
  formatStatementAmount,
  formatStatementDate,
  requestBankStatementImport,
  retargetSheetActions,
} from './bankStatementImportService';

/** The action list the server builds for a two-row import (bank-statement-to-actions.mapper.ts). */
function serverActions(sheetName = 'Bank Statement'): SheetAction[] {
  const headers = ['Date', 'Value Date', 'Description', 'Ref No', 'Debit', 'Credit', 'Balance', 'Source', 'Flag'];
  return [
    { type: 'CREATE_SHEET', sheetName, name: sheetName },
    { type: 'FORMAT_RANGE', sheetName, range: 'C2:D3', format: { numberFormat: '@' } },
    { type: 'FORMAT_RANGE', sheetName, range: 'H2:I3', format: { numberFormat: '@' } },
    {
      type: 'WRITE_TABLE',
      sheetName,
      headers,
      rows: [
        [46123, 46123, 'NEFT IN', 'S1', '', 1627, 2544.04, 'p1 l20', ''],
        [46123, 46124, 'UPI OUT', '004521', 1912, '', 632.04, 'p1 l22', ''],
      ],
    },
    { type: 'FORMAT_RANGE', sheetName, range: 'A1:I1', format: { bold: true } },
    { type: 'FORMAT_RANGE', sheetName, range: 'A2:B3', format: { numberFormat: 'dd-mmm-yyyy' } },
    { type: 'FORMAT_RANGE', sheetName, range: 'E2:G3', format: { numberFormat: '#,##0.00' } },
    ...Array.from({ length: 9 }, (_, col) => ({ type: 'SET_COLUMN_WIDTH', sheetName, col, width: 100 })),
    { type: 'FREEZE_PANES', sheetName, freezeRows: 1 },
  ] as SheetAction[];
}

function result(overrides: Partial<BankStatementImportResult> = {}): BankStatementImportResult {
  return {
    outputSheetName: 'Bank Statement',
    statement: {
      fileName: 'statement.pdf',
      source: 'pdf',
      accountNumberMasked: 'XXXXXXXXXX1234',
      transactionCount: 834,
      periodStart: '2026-04-11',
      periodEnd: '2026-10-08',
    },
    verification: {
      state: 'verified',
      notes: [],
      order: 'oldest_first',
      rowsChecked: 834,
      rowsFailed: 0,
      rowsFlagged: 0,
      directionFixes: 0,
      openingBalance: 950.5,
      openingBalanceSource: 'stated',
      closingBalance: 56506.25,
      totalDebit: 456789.25,
      totalCredit: 512345,
      totalsCheck: 'match',
    },
    exceptions: [],
    skipped: [],
    skippedCount: 0,
    actions: serverActions(),
    ...overrides,
  };
}

describe('the action list the server sends', () => {
  it('is understood in full by the client engine, in the same order', () => {
    const { rich, unsupported } = partitionActions(serverActions());
    expect(unsupported).toEqual([]);
    expect(rich.map((action) => action.type)).toEqual([
      'ADD_SHEET',
      'FORMAT_RANGE',
      'FORMAT_RANGE',
      'WRITE_TABLE',
      'FORMAT_RANGE',
      'FORMAT_RANGE',
      'FORMAT_RANGE',
      ...Array.from({ length: 9 }, () => 'SET_COLUMN_WIDTH'),
      'FREEZE_PANES',
    ]);
  });

  it('keeps the text format ahead of the write and its ranges intact', () => {
    const { rich } = partitionActions(serverActions());
    expect(rich[0]).toMatchObject({ type: 'ADD_SHEET', name: 'Bank Statement' });
    expect(rich[1]).toMatchObject({ sheetName: 'Bank Statement', range: 'C2:D3', format: { numberFormat: '@' } });
    expect(rich[3]).toMatchObject({ type: 'WRITE_TABLE', sheetName: 'Bank Statement' });
  });
});

describe('formatting', () => {
  it('writes dates without going through a time zone', () => {
    expect(formatStatementDate('2026-04-11')).toBe('11 Apr 2026');
    expect(formatStatementDate('2026-01-01')).toBe('1 Jan 2026');
    expect(formatStatementDate(undefined)).toBe('');
    expect(formatStatementDate('11-APR-2026')).toBe('');
  });

  it('writes amounts with Indian digit grouping and two decimals', () => {
    expect(formatStatementAmount(456789.25)).toBe('4,56,789.25');
    expect(formatStatementAmount(512345)).toBe('5,12,345.00');
    expect(formatStatementAmount(-500)).toBe('-500.00');
    expect(formatStatementAmount(undefined)).toBe('');
  });
});

describe('describeVerification', () => {
  it('says every balance ties and the totals match', () => {
    expect(describeVerification(result())).toBe(
      'Checked: all 834 balances tie, and the totals match the statement.',
    );
  });

  it('does not claim a totals match the statement never offered', () => {
    const r = result();
    r.verification.totalsCheck = 'not_stated';
    expect(describeVerification(r)).toBe('Checked: all 834 balances tie.');
  });

  it('counts rows that do not tie apart from rows that could not be read', () => {
    const r = result();
    r.verification = { ...r.verification, state: 'verified_with_exceptions', rowsFailed: 2, rowsFlagged: 3 };
    expect(describeVerification(r)).toBe(
      'Checked with exceptions: 2 rows do not tie, and 1 row could not be read fully. See the Flag column.',
    );
  });

  it('points at the notes when the exception is not about a row', () => {
    const r = result();
    r.verification = { ...r.verification, state: 'verified_with_exceptions', totalsCheck: 'mismatch' };
    expect(describeVerification(r)).toBe('Checked with exceptions. See the notes below.');
  });

  it('never calls an unchecked statement checked', () => {
    const r = result();
    r.verification = {
      ...r.verification,
      state: 'unverified',
      rowsChecked: 0,
      notes: ['The statement has no balance column, so the rows could not be checked.'],
    };
    expect(describeVerification(r)).toBe('The statement has no balance column, so the rows could not be checked.');
    expect(describeVerification(r)).not.toMatch(/^Checked/);
  });
});

describe('buildBankStatementSummary', () => {
  it('describes a verified import', () => {
    expect(buildBankStatementSummary(result())).toEqual({
      contextLine: 'statement.pdf',
      headline: 'Create sheet "Bank Statement" with these transactions?',
      bullets: [
        '834 transactions, 11 Apr 2026 to 8 Oct 2026',
        'Account XXXXXXXXXX1234',
        'Opening balance 950.50, closing balance 56,506.25',
        'Withdrawals 4,56,789.25, deposits 5,12,345.00',
        'Checked: all 834 balances tie, and the totals match the statement.',
      ],
      supportingDetail: 'Nothing is written until you Accept. Reject discards the import.',
    });
  });

  it('carries the notes that explain an exception', () => {
    const r = result();
    r.verification = {
      ...r.verification,
      state: 'verified_with_exceptions',
      rowsFailed: 1,
      rowsFlagged: 1,
      notes: ['1 row could not be fully checked. See the Flag column.'],
    };
    const bullets = buildBankStatementSummary(r).bullets ?? [];
    expect(bullets.slice(-2)).toEqual([
      'Checked with exceptions: 1 row does not tie. See the Flag column.',
      '1 row could not be fully checked. See the Flag column.',
    ]);
  });

  it('does not repeat the reason an import is unverified', () => {
    const note = 'The statement has no balance column, so the rows could not be checked.';
    const r = result();
    r.verification = { ...r.verification, state: 'unverified', notes: [note] };
    delete r.verification.openingBalance;
    delete r.verification.closingBalance;
    const bullets = buildBankStatementSummary(r).bullets ?? [];
    expect(bullets.filter((bullet) => bullet === note)).toHaveLength(1);
    expect(bullets.some((bullet) => bullet.startsWith('Opening balance'))).toBe(false);
  });

  it('leaves out what the statement did not give', () => {
    const r = result();
    r.statement = { fileName: 'one.csv', source: 'csv', transactionCount: 1 };
    expect(buildBankStatementSummary(r).bullets?.[0]).toBe('1 transaction');
  });

  it('says that Accept answers the question that came with the file, without repeating it', () => {
    const question = 'Which are the 10 largest debits?';
    const summary = buildBankStatementSummary(result(), question);
    // The question is already the message above the card.
    expect(JSON.stringify(summary)).not.toContain(question);
    expect(summary.bullets).toEqual(buildBankStatementSummary(result()).bullets);
    expect(summary.supportingDetail).toBe(
      'Accept to create the sheet and I will answer your question. Reject discards both.',
    );
  });

  it('puts the first rows to check and lines left out on the card', () => {
    const r = result({
      exceptions: Array.from({ length: 5 }, (_, i) => ({
        sheetRow: i + 2,
        source: `p1 l${i}`,
        reason: 'Balance does not tie: expected 1.00, statement shows 2.00',
      })),
      skipped: [{ ref: 'p3 l9', text: 'stray text', reason: 'unrecognised' }],
      skippedCount: 1,
    });
    r.verification = { ...r.verification, state: 'verified_with_exceptions', rowsFailed: 5, rowsFlagged: 5 };
    const bullets = buildBankStatementSummary(r).bullets ?? [];
    expect(bullets).toContain('Check sheet row 2 (p1 l0): Balance does not tie: expected 1.00, statement shows 2.00');
    expect(bullets).toContain('Check sheet row 4 (p1 l2): Balance does not tie: expected 1.00, statement shows 2.00');
    expect(bullets.some((bullet) => bullet.startsWith('Check sheet row 5'))).toBe(false);
    expect(bullets).toContain('and 2 more rows, all marked in the Flag column');
    expect(bullets).toContain('Left out (p3 l9): stray text');
  });

  it('contains nothing the card would hide as internal wording', () => {
    const summary = buildBankStatementSummary(result());
    for (const text of [summary.headline, summary.supportingDetail ?? '', ...(summary.bullets ?? [])]) {
      expect(text).not.toMatch(INTERNAL_COPY_MARKERS);
    }
  });
});

describe('buildBankStatementAnswer', () => {
  it('ends by pointing at the card', () => {
    const answer = buildBankStatementAnswer(result());
    expect(answer).toContain('- 834 transactions, 11 Apr 2026 to 8 Oct 2026');
    expect(answer.endsWith('**Accept** to create sheet **Bank Statement**, or **Reject** to discard.')).toBe(true);
    expect(answer).not.toContain('Rows to check');
    expect(answer).not.toContain('Lines left out');
  });

  it('opens with a whole sentence, because the chat shows only that line above the card', () => {
    // TurnRenderer passes the answer through this before showing it over an action card.
    expect(shortenActionPreviewCopy(buildBankStatementAnswer(result()))).toBe(
      'I read statement.pdf and found 834 transactions.',
    );
  });

  it('says the question will be answered after Accept when one came with the file', () => {
    expect(shortenActionPreviewCopy(buildBankStatementAnswer(result(), 'Which are the 10 largest debits?'))).toBe(
      'I read statement.pdf and found 834 transactions. Accept the import below and I will answer your question.',
    );
  });

  it('lists the first flagged rows and counts the rest', () => {
    const r = result({
      exceptions: Array.from({ length: 7 }, (_, i) => ({
        sheetRow: i + 2,
        source: `p1 l${i}`,
        reason: 'Balance does not tie: expected 1.00, statement shows 2.00',
      })),
    });
    r.verification = { ...r.verification, state: 'verified_with_exceptions', rowsFailed: 7, rowsFlagged: 7 };
    const answer = buildBankStatementAnswer(r);
    expect(answer).toContain('- Sheet row 2 (p1 l0): Balance does not tie: expected 1.00, statement shows 2.00');
    expect(answer).toContain('- Sheet row 6 (p1 l4)');
    expect(answer).not.toContain('Sheet row 7');
    expect(answer).toContain('- and 2 more, all marked in the Flag column');
  });

  it('lists lines that were left out', () => {
    const answer = buildBankStatementAnswer(
      result({ skipped: [{ ref: 'p3 l9', text: 'stray text', reason: 'unrecognised' }], skippedCount: 1 }),
    );
    expect(answer).toContain('**Lines left out**\n- p3 l9: stray text');
  });
});

describe('describeBankStatementImportFailure', () => {
  const refusal = (code: string, message = 'server wording') => new BankStatementImportError(message, 422, code);
  const lines = (error: unknown) => describeBankStatementImportFailure(error, 'statement.pdf').split('\n');

  it('names the file, says what was wrong and what to do, one per line', () => {
    expect(lines(refusal('no_header'))).toEqual([
      'Could not import statement.pdf',
      'No bank statement was found in this file. Cellix looks for a table with a Date column and Debit and Credit (or Amount) columns.',
      "Check that this is a statement downloaded from your bank. If it is, this bank's layout is not supported yet.",
    ]);
  });

  it.each(['no_header', 'no_amount_columns', 'no_transactions', 'empty'])(
    'words a %s refusal itself instead of repeating the parser',
    (code) => {
      const [title, reason, nextStep] = lines(refusal(code));
      expect(title).toBe('Could not import statement.pdf');
      expect(reason).toBeTruthy();
      expect(nextStep).toBeTruthy();
      expect([reason, nextStep].join(' ')).not.toMatch(/server wording|transaction table|header row/);
    },
  );

  it("passes on the server's message for a refusal it has no wording for", () => {
    expect(lines(refusal('too_many_transactions', 'This statement has 40000 transactions.'))).toEqual([
      'Could not import statement.pdf',
      'This statement has 40000 transactions.',
    ]);
  });

  it('explains an unreachable server', () => {
    expect(lines(new TypeError('Failed to fetch'))).toEqual([
      'Could not import statement.pdf',
      'The server could not be reached.',
      'Check your connection, then attach the file again.',
    ]);
  });

  it('never shows the text of an unexpected error', () => {
    const text = describeBankStatementImportFailure(new Error("Cannot read properties of undefined (reading 'rows')"), 's.csv');
    expect(text).toBe('Could not import s.csv\nSomething went wrong while reading the statement.\nAttach the file again to retry.');
  });
});

describe('retargetSheetActions', () => {
  it('points every action at the new sheet and nothing else', () => {
    const moved = retargetSheetActions(serverActions(), 'Bank Statement', 'Bank Statement 2');
    expect(moved.every((action) => action.sheetName === 'Bank Statement 2')).toBe(true);
    expect(moved[0]).toMatchObject({ type: 'CREATE_SHEET', name: 'Bank Statement 2' });
    expect(JSON.stringify(moved)).not.toContain('"Bank Statement"');
    // Same rows, untouched.
    expect(moved[3]).toMatchObject({ rows: (serverActions()[3] as { rows: unknown[][] }).rows });
  });

  it('does not change the list it was given', () => {
    const original = serverActions();
    retargetSheetActions(original, 'Bank Statement', 'Bank Statement 2');
    expect(original).toEqual(serverActions());
  });

  it('returns the same list when the name has not changed', () => {
    const original = serverActions();
    expect(retargetSheetActions(original, 'Bank Statement', 'Bank Statement')).toBe(original);
  });
});

describe('findFreeSheetName', () => {
  const existsIn = (names: string[]) => async (candidate: string) => names.includes(candidate);

  it('keeps the name when it is free', async () => {
    expect(await findFreeSheetName('Bank Statement', existsIn(['Sheet1']))).toBe('Bank Statement');
  });

  it('moves to the next free number when it is taken', async () => {
    expect(await findFreeSheetName('Bank Statement', existsIn(['Bank Statement']))).toBe('Bank Statement 2');
    expect(
      await findFreeSheetName('Bank Statement', existsIn(['Bank Statement', 'Bank Statement 2', 'Bank Statement 3'])),
    ).toBe('Bank Statement 4');
  });

  it('counts on from a name that already carries a number', async () => {
    expect(await findFreeSheetName('Bank Statement 2', existsIn(['Bank Statement 2']))).toBe('Bank Statement 3');
    expect(await findFreeSheetName('Bank Statement 2', existsIn(['Bank Statement', 'Bank Statement 2']))).toBe(
      'Bank Statement 3',
    );
  });
});

describe('requestBankStatementImport', () => {
  const rawTable = { source: 'csv' as const, fileName: 's.csv', layout: 'grid' as const, rows: [] };
  const respond = (status: number, body: unknown) =>
    vi.fn(async () => new Response(body === undefined ? 'not json' : JSON.stringify(body), { status }));

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('posts the decoded file with the session cookie and unwraps the envelope', async () => {
    const fetchMock = respond(200, { success: true, data: result() });
    vi.stubGlobal('fetch', fetchMock);

    const received = await requestBankStatementImport({ rawTable, existingSheetNames: ['Sheet1'] });
    expect(received.outputSheetName).toBe('Bank Statement');

    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit];
    expect(url.endsWith('/ingest/bank-statement')).toBe(true);
    expect(init.method).toBe('POST');
    expect(init.credentials).toBe('include');
    expect(JSON.parse(init.body as string)).toEqual({ rawTable, existingSheetNames: ['Sheet1'] });
  });

  it("passes on the server's reason when the file is not a statement", async () => {
    vi.stubGlobal(
      'fetch',
      respond(422, {
        success: false,
        error: { code: 'VALIDATION_ERROR', message: 'Could not find the transaction table.', details: { code: 'no_header' } },
      }),
    );
    await expect(requestBankStatementImport({ rawTable })).rejects.toMatchObject({
      name: 'BankStatementImportError',
      message: 'Could not find the transaction table.',
      status: 422,
      code: 'no_header',
    });
  });

  it.each([
    [401, 'Your session has expired. Sign in again, then retry the import.'],
    [413, 'This statement is too large to import in one go. Split it into shorter periods.'],
    [500, 'The server could not read this statement. Try again in a moment.'],
  ])('explains a %i in plain words', async (status, message) => {
    vi.stubGlobal('fetch', respond(status, { success: false, error: { message: 'internal detail' } }));
    const error = await requestBankStatementImport({ rawTable }).catch((e) => e);
    expect(error).toBeInstanceOf(BankStatementImportError);
    expect(error.message).toBe(message);
  });

  it('copes with an error response that is not JSON', async () => {
    vi.stubGlobal('fetch', respond(502, undefined));
    await expect(requestBankStatementImport({ rawTable })).rejects.toMatchObject({ status: 502 });
  });
});
