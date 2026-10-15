/**
 * Chat-triggered Bank Reconciliation (client side): discover the bank statement + books
 * ledger sheets, run the server match, and hand back the same outcome shapes the GST recon
 * flow uses, so the conversation hook renders it (Accept/Reject card, collision card,
 * docked question, clickable row list) without a second code path.
 */

import {
  BankReconLayout,
  detectBankReconIntent,
  detectBankReconLayout,
  detectBankReconRelayoutTarget,
  isBankReconRelayoutPrompt,
  isCasualBankRecon,
  wantsTimingItemsOnly,
} from '@/utils/bankReconIntent';
import { resolveGstPeriod, ResolvedGstPeriod } from '@/utils/gstPeriod';
import { classifySheetRole, discoverBankSheets, findPriorBrsSheet, readSheetTops, SheetTop } from '@/services/bankSheetDiscovery';
import { readSheetsFull, sheetExists } from '@/services/gstSheetReader';
import { registerGstReconWriteFailureCollision, GstReconChatOutcome } from '@/services/gstReconChat';
import {
  BankReconcileResponse,
  buildBankReconAnswerText,
  buildBankReconUserFacingSummary,
  runBankReconcile,
  runBankRelayout,
} from '@/services/bankReconService';
import { GstReconMissedRow } from '@/types/conversationTurn';
import { RawTable } from '@/types/rawTable';

const PROCESSING_LABEL = 'Bank reconciliation (deterministic match)';
const MAX_LISTED_ROWS = 60;

type Role = 'bank' | 'books';

interface PendingSheetAsk {
  originalMessage: string;
  role: Role;
  candidates: string[];
  chosen: Partial<Record<Role, string>>;
  /** Set when the statement is an attached file (the answer resumes the same attached run). */
  rawTable?: RawTable;
}

interface PendingPeriodAsk {
  originalMessage: string;
  chosen: Partial<Record<Role, string>>;
  rawTable?: RawTable;
}

/** Waiting for "create a sheet from the file, or use the file directly?". */
interface PendingAttachmentChoice {
  message: string;
  rawTable: RawTable;
  /** Sheet this same file was already added as (still present in the workbook). */
  existingSheet?: string | null;
}

// Not worded as a sheet command ("Create sheet…"): the app's local sheet-command lane reacts to that.
const CREATE_OPTION_PREFIX = 'Add as sheet';
const DIRECT_OPTION = 'Use the file directly';

const EXISTING_OPTION_PREFIX = 'Use existing sheet';
const REMEMBER_KEY = 'cellix.bank.addedStatements';
const REMEMBER_MAX = 40;
const OFFER_TTL_MS = 15 * 60 * 1000;

interface AddedStatement {
  sig: string;
  sheet: string;
}

/** Cheap fingerprint of a decoded file: name, size, and a hash of its first and last rows. */
export function statementSignature(table: RawTable): string {
  const text = (r: RawTable['rows'][number]) => r.cells.map((c) => c.t).join('\u241f');
  const edge = [...table.rows.slice(0, 8), ...table.rows.slice(-8)].map(text).join('\u241e');
  let h = 5381;
  for (let i = 0; i < edge.length; i++) h = ((h << 5) + h + edge.charCodeAt(i)) | 0;
  return `${table.fileName}|${table.rows.length}|${(h >>> 0).toString(36)}`;
}

// Survives a reload (localStorage); falls back to memory where storage is unavailable.
let memoryAdded: AddedStatement[] = [];
function loadAdded(): AddedStatement[] {
  try {
    if (typeof localStorage !== 'undefined') {
      const parsed = JSON.parse(localStorage.getItem(REMEMBER_KEY) ?? '[]') as AddedStatement[];
      if (Array.isArray(parsed)) return parsed;
    }
  } catch {
    /* fall through to memory */
  }
  return memoryAdded;
}
function saveAdded(list: AddedStatement[]): void {
  memoryAdded = list;
  try {
    if (typeof localStorage !== 'undefined') localStorage.setItem(REMEMBER_KEY, JSON.stringify(list));
  } catch {
    /* memory copy is enough for this session */
  }
}
function rememberAdded(sig: string, sheet: string): void {
  saveAdded([...loadAdded().filter((e) => e.sheet !== sheet), { sig, sheet }].slice(-REMEMBER_MAX));
}
/** The sheet this same file was already added as, if that sheet is still in the workbook. */
/** Was this sheet written by the "Add as sheet" import (so it must be as trustworthy as the file it came from)? */
function isAddedSheet(sheet: string): boolean {
  return loadAdded().some((e) => e.sheet === sheet);
}

function existingSheetFor(sig: string, sheetNames: string[]): string | null {
  const hit = [...loadAdded()].reverse().find((e) => e.sig === sig && sheetNames.includes(e.sheet));
  return hit?.sheet ?? null;
}

/** An import we asked for: the sheet it creates is recorded once the follow-up reconciliation names it. */
let pendingAddRecord: { sig: string; preferred: string; at: number } | null = null;
const squash = (t: string) => t.toLowerCase().replace(/[^a-z0-9]/g, '');

let pendingAttachmentChoice: PendingAttachmentChoice | null = null;
let pendingSheetAsk: PendingSheetAsk | null = null;
let pendingPeriodAsk: PendingPeriodAsk | null = null;

/** Sheet name → server layout cache id, so "change the layout" can re-map without re-running. */
const layoutCacheIdBySheetName = new Map<string, string>();
let lastSheetName: string | null = null;

export function getBankReconPendingContext(): boolean {
  return pendingSheetAsk !== null || pendingPeriodAsk !== null || pendingAttachmentChoice !== null;
}

export function clearBankReconPendingContext(): void {
  pendingSheetAsk = null;
  pendingPeriodAsk = null;
  pendingAttachmentChoice = null;
  pendingAddRecord = null;
}

const inrShort = (n: number, dir: 'IN' | 'OUT') =>
  `₹${n.toLocaleString('en-IN', { maximumFractionDigits: 2 })} ${dir === 'IN' ? 'in' : 'out'}`;

/** Short source label for a pill: Books / Bank, or the sheet's own name (e.g. "Prior BRS") for any other sheet. */
function sheetLabel(sheet: string, booksName: string, bankName: string): string {
  if (sheet === booksName) return 'Books';
  if (sheet === bankName) return 'Bank';
  return sheet.replace(/\s*\(.*\)\s*$/, '').trim() || sheet;
}

function toListedRows(
  result: BankReconcileResponse,
  layout: BankReconLayout,
  booksName: string,
  bankName: string,
  fileKind?: string,
): GstReconMissedRow[] {
  const kind = fileKind === 'csv' ? 'CSV' : fileKind === 'pdf' ? 'PDF' : fileKind ? 'Excel file' : 'file';
  // List only what the chosen view is about: books_flat → ledger side, bank_flat → bank side, matched → none.
  const side = layout === 'books_flat' ? 'BOOKS' : layout === 'bank_flat' ? 'BANK' : layout === 'matched' ? null : undefined;
  if (side === null) return [];
  return result.rows
    .filter((r) => side === undefined || r.side === side)
    .slice(0, MAX_LISTED_ROWS).map((r) => ({
    vendorName: `${r.side === 'BANK' ? 'Bank' : 'Books'} · ${r.date} · ${(r.narration || r.reference || 'No narration').slice(0, 50)}`,
    gstin: inrShort(r.amount, r.direction),
    // Compact pill: where the row is, nothing else. The details live in the tooltip.
    pillLabel: r.external
      ? `${r.source_ref ?? 'row ' + r.row} in attached ${kind}`
      : `${sheetLabel(r.sheet_name, booksName, bankName)} · row ${r.row}`,
    inFile: r.external ? true : undefined,
    tooltip: r.external
      ? `In the attached ${kind} (${r.sheet_name}), not in your workbook · ${r.source_ref ?? ''} · ${r.date} · ${r.narration || r.reference || 'No narration'} (${inrShort(r.amount, r.direction)})`
      : `${r.side === 'BANK' ? 'Bank statement' : 'Books ledger'} · ${r.date} · ${r.narration || r.reference || 'No narration'} (${inrShort(r.amount, r.direction)}) — click to select row ${r.row}`,
    selectFullRow: true,
    reason: r.reason_label,
    explanation: r.explanation,
    // An attached-file row has no cell in the workbook, so the pill is shown but not clickable.
    sheetName: r.external ? null : r.sheet_name,
    row: r.external ? null : r.row,
  }));
}

/** A sheet name that is safe and short enough to use for the sheet created from an attached file. */
function suggestedSheetName(fileName: string): string {
  const base = fileName.replace(/\.[^.]+$/, '').trim();
  return (base || 'Bank Statement').slice(0, 31);
}

/** Sheets the user named explicitly as @[Sheet] — their role decides which side they are. */
function mentionedSheets(message: string, tops: SheetTop[]): Partial<Record<Role, string>> {
  const out: Partial<Record<Role, string>> = {};
  for (const m of message.matchAll(/@\[([^\]]+)\]/g)) {
    const name = m[1].trim();
    const top = tops.find((t) => t.name === name);
    if (!top) continue;
    const role = classifySheetRole(top);
    if (role === 'bank' && !out.bank) out.bank = name;
    else if (role === 'books' && !out.books) out.books = name;
  }
  return out;
}

function ask(role: Role, candidates: string[], pending: Omit<PendingSheetAsk, 'role' | 'candidates'>): GstReconChatOutcome {
  pendingSheetAsk = { ...pending, role, candidates };
  const what = role === 'bank' ? 'the bank statement' : 'your books bank ledger';
  const question = `Several sheets look like ${what}. Which one should I use?`;
  return { kind: 'question', answer: question, question, options: candidates.slice(0, 6) };
}

export async function tryHandleBankReconChat(
  message: string,
  options?: { conversationId?: string | null },
  carried?: { chosen?: Partial<Record<Role, string>>; period?: ResolvedGstPeriod | null },
): Promise<GstReconChatOutcome | null> {
  if (isBankReconRelayoutPrompt(message)) return handleRelayout(message);

  if (pendingAttachmentChoice) {
    const choice = pendingAttachmentChoice;
    const lower = message.toLowerCase();
    pendingAttachmentChoice = null;
    if (lower.includes(DIRECT_OPTION.toLowerCase())) {
      return tryHandleBankReconWithAttachment(choice.message, choice.rawTable, options, { attachmentMode: 'direct' });
    }
    if (lower.includes(EXISTING_OPTION_PREFIX.toLowerCase()) && choice.existingSheet) {
      // Already added earlier: reconcile against that sheet — no second copy of the same statement.
      return tryHandleBankReconChat(`Reconcile bank ledger against @[${choice.existingSheet}]. ${choice.message}`, options);
    }
    if (lower.includes(CREATE_OPTION_PREFIX.toLowerCase())) {
      const preferredSheetName = suggestedSheetName(choice.rawTable.fileName);
      pendingAddRecord = { sig: statementSignature(choice.rawTable), preferred: preferredSheetName, at: Date.now() };
      return {
        kind: 'import_statement',
        rawTable: choice.rawTable,
        // `{sheet}` becomes the real (unique) name of the sheet the import creates.
        followUpPrompt: 'Reconcile bank ledger against @[{sheet}]',
        preferredSheetName,
      };
    }
    // Unrelated reply — fall through to ordinary handling.
  }

  let chosen: Partial<Record<Role, string>> = { ...(carried?.chosen ?? {}) };
  let forcedPeriod = carried?.period ?? null;

  if (pendingSheetAsk) {
    const ask_ = pendingSheetAsk;
    const hit = ask_.candidates.find((c) => message.toLowerCase().includes(c.toLowerCase()));
    pendingSheetAsk = null;
    if (hit && ask_.rawTable) {
      return tryHandleBankReconWithAttachment(ask_.originalMessage, ask_.rawTable, options, {
        chosenBooks: hit,
        period: forcedPeriod,
        attachmentMode: 'direct',
      });
    }
    if (hit) {
      return tryHandleBankReconChat(ask_.originalMessage, options, {
        chosen: { ...ask_.chosen, [ask_.role]: hit },
        period: forcedPeriod,
      });
    }
    // Unrelated reply — fall through to ordinary intent handling.
  }

  if (pendingPeriodAsk) {
    const p = pendingPeriodAsk;
    const period = resolveGstPeriod(message);
    pendingPeriodAsk = null;
    if (period && p.rawTable) {
      return tryHandleBankReconWithAttachment(p.originalMessage, p.rawTable, options, {
        chosenBooks: p.chosen.books,
        period,
        attachmentMode: 'direct',
      });
    }
    if (period) {
      return tryHandleBankReconChat(p.originalMessage, options, { chosen: p.chosen, period });
    }
  }

  const intent = detectBankReconIntent(message);
  if (!intent) return null;

  if (typeof Excel === 'undefined') {
    return {
      kind: 'message_only',
      answer: 'Bank reconciliation needs an open Excel workbook. Open the file with your bank ledger and bank statement sheets, then ask again.',
    };
  }

  const casual = isCasualBankRecon(message, intent);
  const period = forcedPeriod ?? resolveGstPeriod(message);

  if (!casual && !period) {
    pendingPeriodAsk = { originalMessage: message, chosen };
    return {
      kind: 'message_only',
      answer: 'Which period should I reconcile? For example “April 2025” or “Q1 FY26”.',
    };
  }

  const tops = await readSheetTops();
  if (!tops.length) {
    return { kind: 'message_only', answer: 'This workbook has no sheets I can read. Add your bank ledger and bank statement, then try again.' };
  }
  const found = discoverBankSheets(tops);
  const mentioned = mentionedSheets(message, tops);
  if (pendingAddRecord) {
    const offer = pendingAddRecord;
    // The follow-up of the import we asked for names the sheet it created (it starts with the name we suggested).
    if (mentioned.bank && Date.now() - offer.at < OFFER_TTL_MS && squash(mentioned.bank).startsWith(squash(offer.preferred).slice(0, 20))) {
      rememberAdded(offer.sig, mentioned.bank);
      pendingAddRecord = null;
    } else if (Date.now() - offer.at >= OFFER_TTL_MS) {
      pendingAddRecord = null;
    }
  }
  const bankName = chosen.bank ?? mentioned.bank ?? found.bank;
  const booksName = chosen.books ?? mentioned.books ?? found.books;

  if (!bankName && found.ambiguousBank.length) return ask('bank', found.ambiguousBank, { originalMessage: message, chosen });
  if (!booksName && found.ambiguousBooks.length) return ask('books', found.ambiguousBooks, { originalMessage: message, chosen });

  if (!bankName || !booksName) {
    const missing = [
      !bankName ? '**bank statement** (columns like Date, Narration, Withdrawal/Deposit or Debit/Credit, Balance)' : '',
      !booksName ? '**books bank ledger** (columns like Date, Particulars, Vch Type, Vch No., Debit, Credit)' : '',
    ].filter(Boolean);
    return {
      kind: 'message_only',
      answer:
        `I couldn't find the ${missing.join(' or the ')} in this workbook.\n\n` +
        `Paste or import each into its own sheet (Excel or CSV statements work the same way), then ask again.`,
    };
  }
  chosen = { bank: bankName, books: booksName };

  const grids = await readSheetsFull([bankName, booksName]);
  const bankGrid = grids.find((g) => g.name === bankName);
  const booksGrid = grids.find((g) => g.name === booksName);
  if (!bankGrid?.values?.length || bankGrid.values.length < 2) {
    return { kind: 'message_only', answer: `Sheet "${bankName}" looks empty. Paste the bank statement into it, then ask again.` };
  }
  if (!booksGrid?.values?.length || booksGrid.values.length < 2) {
    return { kind: 'message_only', answer: `Sheet "${booksName}" looks empty. Add the bank ledger rows, then ask again.` };
  }

  const layout: BankReconLayout = detectBankReconLayout(message);

  // Optional: last period's BRS, so its uncleared items carry forward instead of showing up as false mismatches.
  const priorName = findPriorBrsSheet(tops);
  const priorGrid = priorName ? (await readSheetsFull([priorName]))[0] : undefined;

  try {
    const result = await runBankReconcile({
      books_ledger: { sheet_name: booksName, data: booksGrid.values },
      bank_statement: { sheet_name: bankName, data: bankGrid.values, ...(isAddedSheet(bankName) ? { strict: true } : {}) },
      prior_brs: priorGrid?.values?.length ? { sheet_name: priorName!, data: priorGrid.values } : undefined,
      period_start: period?.start,
      period_end: period?.end,
      period_label: period?.label,
      casual,
      layout,
      books_filter: layout === 'books_flat' && wantsTimingItemsOnly(message) ? 'timing' : undefined,
      account_number: intent.extractedAccountNumber,
      client_name: intent.extractedClientName,
      conversation_id: options?.conversationId ?? undefined,
    });

    return buildReconOutcome(result, { booksName, bankName, layout, sheetNames: tops.map((t) => t.name) });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return { kind: 'message_only', answer: `I couldn't complete the bank reconciliation:\n\n${msg}` };
  }
}

/** Turns a server result into the chat outcome (answer text, Accept card, collision, clickable rows). */
async function buildReconOutcome(
  result: BankReconcileResponse,
  ctx: { booksName: string; bankName: string; layout: BankReconLayout; sheetNames: string[] },
): Promise<GstReconChatOutcome> {
  const { booksName, bankName, layout } = ctx;
  {
    if (result.period_zero_message) {
      return { kind: 'message_only', answer: result.period_zero_message };
    }

    const answer = buildBankReconAnswerText(result, booksName, bankName, layout);
    const userFacingSummary = buildBankReconUserFacingSummary(result, booksName, bankName);

    const nothingToWrite = !result.actions?.length;
    if (nothingToWrite) {
      return {
        kind: 'message_only',
        answer: `${answer.split('\n').slice(0, -1).join('\n')}\nNo rows fall under that view, so there's nothing to write.`,
      };
    }

    if (result.layout_cache_id) {
      layoutCacheIdBySheetName.set(result.output_sheet_name, result.layout_cache_id);
      lastSheetName = result.output_sheet_name;
    }

    const taken = ctx.sheetNames.includes(result.output_sheet_name) || (await sheetExists(result.output_sheet_name));
    if (taken) {
      const { collisionId, sheetName } = registerGstReconWriteFailureCollision(result.actions, result.output_sheet_name);
      // The unreconciled-items list belongs to the result, not to the sheet name — keep it on the collision card too.
      return {
        kind: 'sheet_collision',
        answer,
        collisionId,
        sheetName,
        missedRows: toListedRows(result, layout, ctx.booksName, ctx.bankName, result.bank_source?.kind === 'attachment' ? result.bank_source.file_type : undefined),
        missedRowsTitle: 'Unreconciled items',
        missedRowsBeforeAnswer: true,
      };
    }

    return {
      kind: 'recon_ready',
      answer,
      actions: result.actions,
      explanation: userFacingSummary.headline,
      userFacingSummary,
      auditLogId: null,
      sheetName: result.output_sheet_name,
      missedRows: toListedRows(result, layout, ctx.booksName, ctx.bankName, result.bank_source?.kind === 'attachment' ? result.bank_source.file_type : undefined),
      processingLabel: PROCESSING_LABEL,
      missedRowsTitle: 'Unreconciled items',
      missedRowsBeforeAnswer: true,
    };
  }
}

/**
 * Reconcile the employee's own books ledger (a sheet in the workbook) against an ATTACHED
 * statement — CSV, Excel or PDF — read directly. Nothing is written to a sheet for the
 * statement first: the decoded file goes to the server together with the ledger.
 */
export async function tryHandleBankReconWithAttachment(
  message: string,
  rawTable: RawTable,
  options?: { conversationId?: string | null },
  carried?: { chosenBooks?: string; period?: ResolvedGstPeriod | null; attachmentMode?: 'direct' },
): Promise<GstReconChatOutcome> {
  if (isBankReconRelayoutPrompt(message)) return handleRelayout(message);
  if (typeof Excel === 'undefined') {
    return {
      kind: 'message_only',
      answer: 'Reconciling needs an open Excel workbook with your bank ledger in it. Open the file, then attach the statement again.',
    };
  }

  // First ask how the file should be used. A sheet makes every row clickable; the file is still read the
  // same way either way. (Answers resume with attachmentMode 'direct', or go through the import.)
  if (!carried?.attachmentMode) {
    const sheetName = suggestedSheetName(rawTable.fileName);
    const question = 'How should I read the attached statement?';
    // Has this exact file already been added to the workbook? Then offer that sheet instead of a duplicate.
    const sheetNames = (await readSheetTops().catch(() => [] as SheetTop[])).map((t) => t.name);
    const existingSheet = existingSheetFor(statementSignature(rawTable), sheetNames);
    pendingAttachmentChoice = { message, rawTable, existingSheet };
    const lines = [
      question,
      '',
      ...(existingSheet
        ? [`• **${EXISTING_OPTION_PREFIX} “${existingSheet}”** — this file was already added to the workbook as that sheet; reconcile against it (every row stays clickable) without adding a second copy.`]
        : []),
      `• **${CREATE_OPTION_PREFIX} “${sheetName}”** — adds${existingSheet ? ' another copy of' : ''} the statement to your workbook first (you approve it), so every row in the results can be clicked to jump to it.`,
      `• **${DIRECT_OPTION}** — nothing is added to the workbook; rows that are in the file are shown as plain labels (not clickable).`,
    ];
    return {
      kind: 'question',
      answer: lines.join('\n'),
      question,
      options: [
        ...(existingSheet ? [`${EXISTING_OPTION_PREFIX} “${existingSheet}”`] : []),
        `${CREATE_OPTION_PREFIX} “${sheetName}”`,
        DIRECT_OPTION,
      ],
    };
  }

  const intent = detectBankReconIntent(message) ?? { matched: true as const };
  const casual = isCasualBankRecon(message, intent);
  const period = carried?.period ?? resolveGstPeriod(message);
  if (!casual && !period) {
    pendingPeriodAsk = { originalMessage: message, chosen: { books: carried?.chosenBooks }, rawTable };
    return { kind: 'message_only', answer: 'Which period should I reconcile? For example “April 2025” or “Q1 FY26”.' };
  }

  const tops = await readSheetTops();
  if (!tops.length) {
    return { kind: 'message_only', answer: 'This workbook has no sheets I can read. Add your bank ledger sheet, then attach the statement again.' };
  }
  const found = discoverBankSheets(tops);

  // The statement is the attached file, so the ledger is whichever sheet looks like one. A sheet that
  // reads as a bank statement is not the ledger; an unclassified transaction table is.
  const mentioned = mentionedSheets(message, tops);
  let booksName = carried?.chosenBooks ?? mentioned.books ?? found.books ?? undefined;
  if (!booksName) {
    const pool = (found.ambiguousBooks.length ? found.ambiguousBooks : found.candidates).filter((n) => n !== found.bank);
    if (pool.length === 1) booksName = pool[0];
    else if (pool.length > 1) return ask('books', pool, { originalMessage: message, chosen: {}, rawTable });
  }
  if (!booksName) {
    return {
      kind: 'message_only',
      answer:
        'I couldn\'t find your **books bank ledger** in this workbook (a sheet with columns like Date, Particulars, Debit, Credit or Receipts, Payments).\n\n' +
        'Open the workbook that has the ledger, then attach the statement again.',
    };
  }

  const sheets = await readSheetsFull([booksName]);
  const booksGrid = sheets.find((g) => g.name === booksName);
  if (!booksGrid?.values?.length || booksGrid.values.length < 2) {
    return { kind: 'message_only', answer: `Sheet "${booksName}" looks empty. Add the bank ledger rows, then try again.` };
  }
  const priorName = findPriorBrsSheet(tops);
  const priorGrid = priorName ? (await readSheetsFull([priorName]))[0] : undefined;

  const layout: BankReconLayout = detectBankReconLayout(message);
  try {
    const result = await runBankReconcile({
      books_ledger: { sheet_name: booksName, data: booksGrid.values },
      bank_statement: { sheet_name: rawTable.fileName, raw_table: rawTable },
      prior_brs: priorGrid?.values?.length ? { sheet_name: priorName!, data: priorGrid.values } : undefined,
      period_start: period?.start,
      period_end: period?.end,
      period_label: period?.label,
      casual,
      layout,
      books_filter: layout === 'books_flat' && wantsTimingItemsOnly(message) ? 'timing' : undefined,
      account_number: intent.extractedAccountNumber,
      client_name: intent.extractedClientName,
      conversation_id: options?.conversationId ?? undefined,
    });
    return buildReconOutcome(result, { booksName, bankName: rawTable.fileName, layout, sheetNames: tops.map((t) => t.name) });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return { kind: 'message_only', answer: `I couldn't complete the bank reconciliation:\n\n${msg}` };
  }
}

const LAYOUT_LABELS: Record<BankReconLayout, string> = {
  categorized: 'the summary grouped by reason',
  books_flat: 'a single table of ledger entries missing from the bank',
  bank_flat: 'a single table of bank entries to post',
  matched: 'the matched list',
};

async function handleRelayout(message: string): Promise<GstReconChatOutcome> {
  const cacheId = (lastSheetName && layoutCacheIdBySheetName.get(lastSheetName)) || [...layoutCacheIdBySheetName.values()].pop();
  if (!cacheId) {
    return {
      kind: 'message_only',
      answer: "I don't have a recent bank reconciliation to re-lay-out — run it again and I can change the layout right after.",
    };
  }
  const layout = detectBankReconRelayoutTarget(message);
  try {
    const out = await runBankRelayout({ layout_cache_id: cacheId, layout });
    layoutCacheIdBySheetName.set(out.output_sheet_name, cacheId);
    lastSheetName = out.output_sheet_name;
    const answer = `Switched to ${LAYOUT_LABELS[layout]} — same rows, same counts, just re-formatted. It writes to its own sheet: "${out.output_sheet_name}".`;
    if (!out.actions.length) {
      return { kind: 'message_only', answer: `${answer.split(' It writes')[0]} But no rows fall under that view, so there's nothing to write.` };
    }
    if (await sheetExists(out.output_sheet_name)) {
      const { collisionId, sheetName } = registerGstReconWriteFailureCollision(out.actions, out.output_sheet_name);
      return { kind: 'sheet_collision', answer, collisionId, sheetName };
    }
    return {
      kind: 'recon_ready',
      answer,
      actions: out.actions,
      explanation: answer,
      userFacingSummary: { headline: answer, bullets: [] },
      sheetName: out.output_sheet_name,
      processingLabel: PROCESSING_LABEL,
    };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return { kind: 'message_only', answer: `I couldn't change the layout: ${msg}` };
  }
}
