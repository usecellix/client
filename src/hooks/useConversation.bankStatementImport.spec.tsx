// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useConversation } from '@/hooks/useConversation';
import {
  BankStatementImportError,
  BankStatementImportResult,
  requestBankStatementImport,
} from '@/services/bankStatementImportService';
import type { WorkbookContext } from '@/types/cellix.types';
import type { ActionBlock } from '@/types/conversationTurn';
import type { RawTable } from '@/types/rawTable';
import type { SheetAction } from '@/types/sheet-actions';

/**
 * Attaching a statement and asking about it (ATTACHMENT_EXTRACTION_PLAN.md
 * §6.5): the import is staged for Accept, and the question typed alongside the
 * file is asked as an ordinary chat turn once the statement is in the workbook.
 */

const workbook = vi.hoisted(() => ({ sheets: ['Sheet1'] as string[] }));

vi.mock('@/utils/missingSheetRetry', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@/utils/missingSheetRetry')>()),
  listWorksheetNames: vi.fn(async () => workbook.sheets),
}));
vi.mock('@/services/gstSheetReader', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@/services/gstSheetReader')>()),
  sheetExists: vi.fn(async (name: string) => workbook.sheets.includes(name)),
}));
vi.mock('@/services/bankStatementImportService', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@/services/bankStatementImportService')>()),
  requestBankStatementImport: vi.fn(),
}));
vi.mock('@/utils/pendingWorkbookContext', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@/utils/pendingWorkbookContext')>()),
  getContextForSend: vi.fn(async () => ({
    sheetData: [['Date', 'Debit'], [46123, 1912]],
    workbookContext: { activeSheet: 'Bank Statement', sheets: ['Sheet1', 'Bank Statement'] } as unknown as WorkbookContext,
    promptContext: 'Bank Statement has 834 rows.',
  })),
}));

const request = vi.mocked(requestBankStatementImport);

const RAW_TABLE: RawTable = { source: 'pdf', fileName: 'statement.pdf', layout: 'positioned', rows: [] };
const QUESTION = 'Which are the 10 largest debits in Bank Statement?';

const ACTIONS = [
  { type: 'CREATE_SHEET', sheetName: 'Bank Statement', name: 'Bank Statement' },
  { type: 'WRITE_TABLE', sheetName: 'Bank Statement', headers: ['Date', 'Debit'], rows: [[46123, 1912]] },
  { type: 'FREEZE_PANES', sheetName: 'Bank Statement', freezeRows: 1 },
] as SheetAction[];

function importResult(): BankStatementImportResult {
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
    actions: ACTIONS,
  };
}

function sseBlock(type: string, data: unknown): string {
  return `event: ${type}\ndata: ${JSON.stringify(data)}\n\n`;
}

function chatReply(text: string): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(encoder.encode(sseBlock('chunk', { text })));
      controller.enqueue(encoder.encode(sseBlock('conversation_end', { summary: text })));
      controller.close();
    },
  });
  return { ok: true, status: 200, body: stream, headers: new Headers() } as unknown as Response;
}

/** The chat requests the hook sent, by the message each one carried. */
function chatMessages(fetchMock: ReturnType<typeof vi.fn>): string[] {
  return fetchMock.mock.calls
    .filter(([url]) => String(url).includes('/excel-ai/conversation'))
    .map(([, init]) => (JSON.parse((init as RequestInit).body as string) as { message: string }).message);
}

describe('useConversation — importing an attached statement', () => {
  let fetchMock: ReturnType<typeof vi.fn>;
  let onActions: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    localStorage.clear();
    workbook.sheets = ['Sheet1'];
    request.mockReset();
    request.mockResolvedValue(importResult());
    fetchMock = vi.fn(async () => chatReply('The largest debit is 30,000.00 on 5 May 2026.'));
    vi.stubGlobal('fetch', fetchMock);
    onActions = vi.fn().mockResolvedValue(undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  async function runImport(workbookKey: string, followUpPrompt?: string) {
    const onClearPreview = vi.fn();
    const { result } = renderHook(() => useConversation({ workbookKey, onActions, onClearPreview }));
    let ok = false;
    await act(async () => {
      ok = await result.current.importBankStatement(RAW_TABLE, 'statement.pdf', { mode: 'ask', followUpPrompt });
    });
    await waitFor(() => expect(result.current.turns).toHaveLength(1));
    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions');
    return { result, ok, turn, block, onClearPreview };
  }

  it('stages the import as a card, with no model call', async () => {
    const { ok, turn, block } = await runImport('import-plain.xlsx');

    expect(ok).toBe(true);
    // The thread shows the file as an attachment on the message, not an internal label.
    expect(turn.userMessage).toBe('Import this bank statement');
    expect(turn.attachmentNames).toEqual(['statement.pdf']);
    expect(block).toMatchObject({
      proposalStatus: 'pending',
      importSheetName: 'Bank Statement',
      actions: ACTIONS,
    });
    expect(block?.followUp).toBeUndefined();
    expect(request).toHaveBeenCalledWith(
      { rawTable: RAW_TABLE, existingSheetNames: ['Sheet1'] },
      expect.anything(),
    );
    expect(chatMessages(fetchMock)).toEqual([]);
    expect(onActions).not.toHaveBeenCalled();
  });

  it('keeps the statement figures out of the history that later prompts send to the model', async () => {
    const { result } = await runImport('import-history.xlsx');
    const history = JSON.stringify(result.current.sessions[0].history);
    expect(history).toContain('statement.pdf');
    expect(history).not.toMatch(/950\.50|56,506|4,56,789|XXXXXXXXXX1234/);
  });

  it('carries the question typed with the file on the card, and does not ask it yet', async () => {
    const { turn, block } = await runImport('import-question.xlsx', QUESTION);

    // The message in the thread is the question itself, with the file attached.
    expect(turn.userMessage).toBe(QUESTION);
    expect(turn.attachmentNames).toEqual(['statement.pdf']);
    expect(block?.followUp).toEqual({ prompt: QUESTION, mode: 'ask' });
    expect(chatMessages(fetchMock)).toEqual([]);
  });

  it('asks the question as an ordinary chat turn once the import is accepted', async () => {
    const { result, turn, block } = await runImport('import-accept.xlsx', QUESTION);

    await act(async () => {
      await result.current.acceptActions(turn.id, block!.id);
    });
    expect(onActions).toHaveBeenCalledTimes(1);
    expect(onActions.mock.calls[0][0]).toEqual(ACTIONS);

    await waitFor(() => expect(chatMessages(fetchMock)).toEqual([QUESTION]));
    // One exchange, not two messages: the answer arrives in a turn that carries
    // on the import's turn and is drawn without a user message of its own.
    await waitFor(() => expect(result.current.turns).toHaveLength(2));
    const [importTurn, answerTurn] = result.current.turns;
    expect(importTurn.userMessage).toBe(QUESTION);
    expect(importTurn.continuesTurnId).toBeUndefined();
    expect(answerTurn.continuesTurnId).toBe(importTurn.id);
    expect(answerTurn.attachmentNames).toBeUndefined();
    // It was sent with a fresh read of the workbook, in the mode it was typed in.
    const body = JSON.parse((fetchMock.mock.calls[0][1] as RequestInit).body as string) as { mode?: string };
    expect(body.mode).toBe('ask');
  });

  it('asks nothing after Accept when no question came with the file', async () => {
    const { result, turn, block } = await runImport('import-noquestion.xlsx');
    await act(async () => {
      await result.current.acceptActions(turn.id, block!.id);
    });
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(chatMessages(fetchMock)).toEqual([]);
    expect(result.current.turns).toHaveLength(1);
  });

  it('drops the question when the import is rejected', async () => {
    const { result, turn, block } = await runImport('import-reject.xlsx', QUESTION);
    await act(async () => {
      await result.current.rejectActions(turn.id, block!.id);
    });
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(onActions).not.toHaveBeenCalled();
    expect(chatMessages(fetchMock)).toEqual([]);
  });

  it('does not ask the question when Accept fails', async () => {
    onActions.mockRejectedValueOnce(new Error('WRITE_TABLE: boom'));
    const { result, turn, block } = await runImport('import-acceptfail.xlsx', QUESTION);
    await act(async () => {
      await result.current.acceptActions(turn.id, block!.id).catch(() => undefined);
    });
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(chatMessages(fetchMock)).toEqual([]);
  });

  it('moves to a free sheet name if the one it was built for has been taken by Accept time', async () => {
    const { result, turn, block } = await runImport('import-collision.xlsx');
    workbook.sheets = ['Sheet1', 'Bank Statement'];

    await act(async () => {
      await result.current.acceptActions(turn.id, block!.id);
    });
    const applied = onActions.mock.calls[0][0] as SheetAction[];
    expect(applied.map((action) => action.sheetName)).toEqual([
      'Bank Statement 2',
      'Bank Statement 2',
      'Bank Statement 2',
    ]);
    const after = result.current.turns[0].blocks.find((b): b is ActionBlock => b.type === 'actions');
    expect(after).toMatchObject({ proposalStatus: 'accepted', importSheetName: 'Bank Statement 2' });
    expect(after?.userFacingSummary?.headline).toBe('Create sheet "Bank Statement 2" with these transactions?');
  });

  it('reports a refused file on the turn and says the import failed', async () => {
    request.mockRejectedValueOnce(new BankStatementImportError('server wording', 422, 'no_header'));
    const { ok, turn, block } = await runImport('import-refused.xlsx', QUESTION);

    expect(ok).toBe(false);
    expect(block).toBeUndefined();
    expect(turn.phase).toBe('error');
    expect(turn.error?.split('\n')[0]).toBe('Could not import statement.pdf');
    expect(turn.error).toContain('No bank statement was found in this file.');
    expect(chatMessages(fetchMock)).toEqual([]);
  });
});
