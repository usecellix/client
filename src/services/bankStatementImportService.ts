/**
 * Client API for POST /ingest/bank-statement, plus the wording of the result.
 * Root/ATTACHMENT_EXTRACTION_PLAN.md §6.
 */

import { getBankStatementImportEndpoint } from '@/lib/apiConfig';
import { RawTable, RawTableSource } from '@/types/rawTable';
import { SheetAction } from '@/types/sheet-actions';
import { UserFacingSummary } from '@/utils/userFacingResponse';

export type BankStatementVerificationState = 'verified' | 'verified_with_exceptions' | 'unverified';

export interface BankStatementVerification {
  state: BankStatementVerificationState;
  notes: string[];
  order: 'oldest_first' | 'newest_first';
  rowsChecked: number;
  rowsFailed: number;
  rowsFlagged: number;
  directionFixes: number;
  openingBalance?: number;
  openingBalanceSource: 'stated' | 'derived' | 'none';
  closingBalance?: number;
  totalDebit: number;
  totalCredit: number;
  totalsCheck: 'match' | 'mismatch' | 'not_stated';
}

export interface BankStatementImportResult {
  outputSheetName: string;
  statement: {
    fileName: string;
    source: RawTableSource;
    accountNumberMasked?: string;
    transactionCount: number;
    periodStart?: string;
    periodEnd?: string;
  };
  verification: BankStatementVerification;
  exceptions: Array<{ sheetRow: number; source: string; reason: string }>;
  skipped: Array<{ ref: string; text: string; reason: 'unrecognised' | 'after_total' }>;
  skippedCount: number;
  actions: SheetAction[];
}

export interface BankStatementImportRequest {
  rawTable: RawTable;
  existingSheetNames?: string[];
  activeSheetName?: string;
  /** Name the new sheet (e.g. after the attached file). */
  preferredSheetName?: string;
}

/** A refusal from the server, with a message written for the user. */
export class BankStatementImportError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code?: string,
  ) {
    super(message);
    this.name = 'BankStatementImportError';
  }
}

function messageForStatus(status: number, serverMessage: string | undefined): string {
  if (status === 401) return 'Your session has expired. Sign in again, then retry the import.';
  if (status === 413) return 'This statement is too large to import in one go. Split it into shorter periods.';
  // 422 carries a reason the user can act on (not a statement, no transactions…).
  if (status === 422 && serverMessage) return serverMessage;
  if (status >= 500) return 'The server could not read this statement. Try again in a moment.';
  return serverMessage || `The import failed (${status}).`;
}

export async function requestBankStatementImport(
  request: BankStatementImportRequest,
  signal?: AbortSignal,
): Promise<BankStatementImportResult> {
  const url = getBankStatementImportEndpoint();
  const headers: Record<string, string> = { 'Content-Type': 'application/json' };
  if (url.includes('.ngrok-free.app')) {
    headers['ngrok-skip-browser-warning'] = 'true';
  }

  const response = await fetch(url, {
    method: 'POST',
    credentials: 'include',
    headers,
    body: JSON.stringify(request),
    signal,
  });

  const body: unknown = await response.json().catch(() => null);
  if (!response.ok) {
    const error = (body as { error?: { message?: string; details?: { code?: string } } } | null)?.error;
    throw new BankStatementImportError(
      messageForStatus(response.status, error?.message),
      response.status,
      error?.details?.code,
    );
  }
  if (body && typeof body === 'object' && 'data' in body) {
    return (body as { data: BankStatementImportResult }).data;
  }
  return body as BankStatementImportResult;
}

/**
 * What the server found wrong with the file, in the user's terms. The server's
 * own message names parser concepts ("transaction table", "header row"); these
 * say what Cellix was looking for and what to do next.
 */
const REFUSAL_REASONS: Record<string, { reason: string; nextStep: string }> = {
  no_header: {
    reason:
      'No bank statement was found in this file. Cellix looks for a table with a Date column and Debit and Credit (or Amount) columns.',
    nextStep:
      'Check that this is a statement downloaded from your bank. If it is, this bank\'s layout is not supported yet.',
  },
  no_amount_columns: {
    reason: 'A table with a Date column was found, but it has no Debit, Credit or Amount column.',
    nextStep: 'Download the statement again with the transaction amounts included.',
  },
  no_transactions: {
    reason: 'The statement\'s column headings were found, but there are no transactions under them.',
    nextStep: 'Check that the statement covers a period with transactions.',
  },
  empty: {
    reason: 'There is nothing readable in this file.',
    nextStep: 'Download the statement again from your bank and attach the new copy.',
  },
};

/**
 * The message shown in the chat when an import fails: the file, the reason and
 * what to do, one per line. The first line is the title of the error card.
 */
export function describeBankStatementImportFailure(error: unknown, fileName: string): string {
  const title = `Could not import ${fileName}`;
  if (error instanceof BankStatementImportError) {
    const known = error.code ? REFUSAL_REASONS[error.code] : undefined;
    if (known) return [title, known.reason, known.nextStep].join('\n');
    return [title, error.message].join('\n');
  }
  if (error instanceof Error && error.message === 'Failed to fetch') {
    return [title, 'The server could not be reached.', 'Check your connection, then attach the file again.'].join('\n');
  }
  return [title, 'Something went wrong while reading the statement.', 'Attach the file again to retry.'].join('\n');
}

// ── Wording ─────────────────────────────────────────────────────────────────

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** "2026-04-11" -> "11 Apr 2026". Formatted by hand: `new Date(iso)` shifts the day in western time zones. */
export function formatStatementDate(iso: string | undefined): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso ?? '');
  if (!match) return '';
  return `${Number(match[3])} ${MONTHS[Number(match[2]) - 1]} ${match[1]}`;
}

const amountFormat = new Intl.NumberFormat('en-IN', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

export function formatStatementAmount(amount: number | undefined): string {
  return amount === undefined ? '' : amountFormat.format(amount);
}

const plural = (count: number, one: string, many: string) => `${count} ${count === 1 ? one : many}`;

/** One sentence saying how far the import was checked. */
export function describeVerification(result: BankStatementImportResult): string {
  const v = result.verification;
  if (v.state === 'unverified') {
    return v.notes[0] ?? 'The rows could not be checked.';
  }
  if (v.state === 'verified') {
    const totals = v.totalsCheck === 'match' ? ', and the totals match the statement' : '';
    return `Checked: all ${plural(v.rowsChecked, 'balance ties', 'balances tie')}${totals}.`;
  }
  const failed = v.rowsFailed > 0 ? `${plural(v.rowsFailed, 'row does', 'rows do')} not tie` : '';
  const flaggedOnly = v.rowsFlagged - v.rowsFailed;
  const flagged = flaggedOnly > 0 ? `${plural(flaggedOnly, 'row', 'rows')} could not be read fully` : '';
  const parts = [failed, flagged].filter(Boolean);
  return parts.length
    ? `Checked with exceptions: ${parts.join(', and ')}. See the Flag column.`
    : 'Checked with exceptions. See the notes below.';
}

function periodText(result: BankStatementImportResult): string {
  const start = formatStatementDate(result.statement.periodStart);
  const end = formatStatementDate(result.statement.periodEnd);
  if (!start || !end) return '';
  return start === end ? start : `${start} to ${end}`;
}

/** Facts about the statement, shared by the chat answer and the card. */
function statementBullets(result: BankStatementImportResult): string[] {
  const v = result.verification;
  const period = periodText(result);
  const bullets = [
    `${plural(result.statement.transactionCount, 'transaction', 'transactions')}${period ? `, ${period}` : ''}`,
  ];
  if (result.statement.accountNumberMasked) {
    bullets.push(`Account ${result.statement.accountNumberMasked}`);
  }
  if (v.openingBalance !== undefined && v.closingBalance !== undefined) {
    bullets.push(
      `Opening balance ${formatStatementAmount(v.openingBalance)}, closing balance ${formatStatementAmount(v.closingBalance)}`,
    );
  }
  bullets.push(
    `Withdrawals ${formatStatementAmount(v.totalDebit)}, deposits ${formatStatementAmount(v.totalCredit)}`,
  );
  bullets.push(describeVerification(result));
  // Notes explain anything short of verified. The first note of an unverified
  // result is already the verification sentence above.
  const notes = v.state === 'unverified' ? v.notes.slice(1) : v.notes;
  bullets.push(...notes);
  return bullets;
}

const MAX_LISTED_ON_CARD = 3;
const clipText = (text: string, max: number) => (text.length > max ? `${text.slice(0, max - 1).trimEnd()}…` : text);

/** Rows to check and lines left out, short enough to sit on the card. */
function problemBullets(result: BankStatementImportResult): string[] {
  const bullets: string[] = [];
  const listedRows = result.exceptions.slice(0, MAX_LISTED_ON_CARD);
  for (const exception of listedRows) {
    bullets.push(`Check sheet row ${exception.sheetRow} (${exception.source}): ${exception.reason}`);
  }
  const moreRows = result.verification.rowsFlagged - listedRows.length;
  if (listedRows.length > 0 && moreRows > 0) {
    bullets.push(`and ${plural(moreRows, 'more row', 'more rows')}, all marked in the Flag column`);
  }
  const listedSkipped = result.skipped.slice(0, MAX_LISTED_ON_CARD);
  for (const skipped of listedSkipped) {
    bullets.push(`Left out (${skipped.ref}): ${clipText(skipped.text, 70)}`);
  }
  const moreSkipped = result.skippedCount - listedSkipped.length;
  if (listedSkipped.length > 0 && moreSkipped > 0) {
    bullets.push(`and ${plural(moreSkipped, 'more line', 'more lines')} left out`);
  }
  return bullets;
}

/**
 * The card under the answer. It carries everything about the import, because
 * the chat shows only the first sentence of the answer above a card.
 * `followUpPrompt` is the question the user typed alongside the file.
 */
export function buildBankStatementSummary(
  result: BankStatementImportResult,
  followUpPrompt?: string,
): UserFacingSummary {
  const bullets = [...statementBullets(result), ...problemBullets(result)];
  return {
    contextLine: result.statement.fileName,
    headline: `Create sheet "${result.outputSheetName}" with these transactions?`,
    bullets,
    supportingDetail: followUpPrompt
      ? 'Accept to create the sheet and I will answer your question. Reject discards both.'
      : 'Nothing is written until you Accept. Reject discards the import.',
  };
}

const MAX_EXCEPTIONS_IN_ANSWER = 5;

export function buildBankStatementAnswer(result: BankStatementImportResult, followUpPrompt?: string): string {
  const count = plural(result.statement.transactionCount, 'transaction', 'transactions');
  const lines = [
    // The chat shows only this first line above the card, so it is a whole sentence.
    followUpPrompt
      ? `I read ${result.statement.fileName} and found ${count}. Accept the import below and I will answer your question.`
      : `I read ${result.statement.fileName} and found ${count}.`,
    '',
    ...statementBullets(result).map((bullet) => `- ${bullet}`),
  ];

  if (result.exceptions.length) {
    lines.push('', '**Rows to check**');
    for (const exception of result.exceptions.slice(0, MAX_EXCEPTIONS_IN_ANSWER)) {
      lines.push(`- Sheet row ${exception.sheetRow} (${exception.source}): ${exception.reason}`);
    }
    const more = result.verification.rowsFlagged - Math.min(result.exceptions.length, MAX_EXCEPTIONS_IN_ANSWER);
    if (more > 0) lines.push(`- and ${more} more, all marked in the Flag column`);
  }

  if (result.skipped.length) {
    lines.push('', '**Lines left out**');
    for (const skipped of result.skipped.slice(0, MAX_EXCEPTIONS_IN_ANSWER)) {
      lines.push(`- ${skipped.ref}: ${skipped.text}`);
    }
    const more = result.skippedCount - Math.min(result.skipped.length, MAX_EXCEPTIONS_IN_ANSWER);
    if (more > 0) lines.push(`- and ${more} more`);
  }

  lines.push(
    '',
    `Review the card below and **Accept** to create sheet **${result.outputSheetName}**, or **Reject** to discard.`,
  );
  return lines.join('\n');
}

// ── Sheet name at Accept time ───────────────────────────────────────────────

/**
 * Points every action at `to` instead of `from`. Used when the name the import
 * was built for has been taken between building the card and accepting it: the
 * import moves to a free name, it never writes into the sheet that now exists.
 */
export function retargetSheetActions(actions: SheetAction[], from: string, to: string): SheetAction[] {
  if (from === to) return actions;
  return actions.map((action) => {
    const next = { ...action } as SheetAction & { name?: string };
    if (next.sheetName === from) next.sheetName = to;
    if (next.name === from) next.name = to;
    return next;
  });
}

const MAX_NAME_ATTEMPTS = 50;

/**
 * `name` if no sheet has it, else the first free "name 2", "name 3"…
 * `exists` is injected so this stays testable without Excel.
 */
export async function findFreeSheetName(
  name: string,
  exists: (candidate: string) => Promise<boolean>,
): Promise<string> {
  if (!(await exists(name))) return name;
  const base = name.replace(/ \d+$/, '');
  for (let n = 2; n < MAX_NAME_ATTEMPTS; n++) {
    const candidate = `${base} ${n}`;
    if (!(await exists(candidate))) return candidate;
  }
  throw new Error(`No free sheet name found for "${name}".`);
}
