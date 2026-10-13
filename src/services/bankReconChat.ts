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
import { discoverBankSheets, findPriorBrsSheet, readSheetTops } from '@/services/bankSheetDiscovery';
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

const PROCESSING_LABEL = 'Bank reconciliation (deterministic match)';
const MAX_LISTED_ROWS = 60;

type Role = 'bank' | 'books';

interface PendingSheetAsk {
  originalMessage: string;
  role: Role;
  candidates: string[];
  chosen: Partial<Record<Role, string>>;
}

interface PendingPeriodAsk {
  originalMessage: string;
  chosen: Partial<Record<Role, string>>;
}

let pendingSheetAsk: PendingSheetAsk | null = null;
let pendingPeriodAsk: PendingPeriodAsk | null = null;

/** Sheet name → server layout cache id, so "change the layout" can re-map without re-running. */
const layoutCacheIdBySheetName = new Map<string, string>();
let lastSheetName: string | null = null;

export function getBankReconPendingContext(): boolean {
  return pendingSheetAsk !== null || pendingPeriodAsk !== null;
}

export function clearBankReconPendingContext(): void {
  pendingSheetAsk = null;
  pendingPeriodAsk = null;
}

const inrShort = (n: number, dir: 'IN' | 'OUT') =>
  `₹${n.toLocaleString('en-IN', { maximumFractionDigits: 2 })} ${dir === 'IN' ? 'in' : 'out'}`;

function toListedRows(result: BankReconcileResponse, layout: BankReconLayout): GstReconMissedRow[] {
  // List only what the chosen view is about: books_flat → ledger side, bank_flat → bank side, matched → none.
  const side = layout === 'books_flat' ? 'BOOKS' : layout === 'bank_flat' ? 'BANK' : layout === 'matched' ? null : undefined;
  if (side === null) return [];
  return result.rows
    .filter((r) => side === undefined || r.side === side)
    .slice(0, MAX_LISTED_ROWS).map((r) => ({
    vendorName: `${r.side === 'BANK' ? 'Bank' : 'Books'} · ${r.date} · ${(r.narration || r.reference || 'No narration').slice(0, 50)}`,
    gstin: inrShort(r.amount, r.direction),
    reason: r.reason_label,
    explanation: r.explanation,
    sheetName: r.sheet_name,
    row: r.row,
  }));
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

  let chosen: Partial<Record<Role, string>> = { ...(carried?.chosen ?? {}) };
  let forcedPeriod = carried?.period ?? null;

  if (pendingSheetAsk) {
    const ask_ = pendingSheetAsk;
    const hit = ask_.candidates.find((c) => message.toLowerCase().includes(c.toLowerCase()));
    pendingSheetAsk = null;
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
  const bankName = chosen.bank ?? found.bank;
  const booksName = chosen.books ?? found.books;

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
      bank_statement: { sheet_name: bankName, data: bankGrid.values },
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

    const taken = tops.some((t) => t.name === result.output_sheet_name) || (await sheetExists(result.output_sheet_name));
    if (taken) {
      const { collisionId, sheetName } = registerGstReconWriteFailureCollision(result.actions, result.output_sheet_name);
      return { kind: 'sheet_collision', answer, collisionId, sheetName };
    }

    return {
      kind: 'recon_ready',
      answer,
      actions: result.actions,
      explanation: userFacingSummary.headline,
      userFacingSummary,
      auditLogId: null,
      sheetName: result.output_sheet_name,
      missedRows: toListedRows(result, layout),
      processingLabel: PROCESSING_LABEL,
      missedRowsTitle: 'Unreconciled items',
    };
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
