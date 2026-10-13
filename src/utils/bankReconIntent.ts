/**
 * Detect Bank Reconciliation (books bank ledger vs. bank statement / BRS) intents from chat.
 * Deliberately never claims a GST prompt — GST recon owns anything that mentions GSTR / 2B / ITC.
 */

export type BankReconLayout = 'categorized' | 'books_flat' | 'bank_flat' | 'matched';

export interface BankReconIntent {
  matched: true;
  /** Account number quoted in the prompt (9–18 digits) — makes the run "gated". */
  extractedAccountNumber?: string;
  extractedClientName?: string;
}

const RECON_VERB =
  /\b(reconcil(?:e|iation|ing)?|recon|match(?:ing)?|compare|cross[-\s]?check|tally|tie)\b/i;
const BANK_MARKERS = /\b(bank|brs|passbook|bank\s*book|cash\s*book)\b/i;
const GST_MARKERS = /\b(gst|gstr|gstr[-\s]?2[ab]|gstr[-\s]?1|2b|2a|itc|ims|purchase\s+register|sales\s+register)\b/i;
const BRS_PHRASE = /\bbrs\b|\bbank\s+reconciliation(?:\s+statement)?\b|\bbank\s+recon\b/i;
const ACCOUNT_RE = /\b(?:a\/c|acc(?:oun)?t)(?:\s*(?:no\.?|number|#))?\s*[:\-]?\s*(\d{9,18})\b/i;

const RELAYOUT_VERB = /\b(change|switch|convert|turn|redo|rearrange|rebuild|reformat|re-?layout|make|put|show|give)\b/i;
const RELAYOUT_WORDS =
  /\b(layout|single\s+table|one\s+table|flat\s+table|flat\s+list|categorized|sections?|grouped?|by\s+reason|books[-\s]?only|bank[-\s]?only|matched)\b/i;

export function detectBankReconIntent(message: string): BankReconIntent | null {
  const text = message.trim();
  if (!text) return null;
  if (GST_MARKERS.test(text)) return null;

  const hasBank = BANK_MARKERS.test(text);
  const brs = BRS_PHRASE.test(text);
  if (!brs && !(hasBank && RECON_VERB.test(text))) return null;
  // "bank" inside an unrelated sentence ("bank holiday") with no recon verb is excluded above.

  const out: BankReconIntent = { matched: true };
  const acct = text.match(ACCOUNT_RE);
  if (acct) out.extractedAccountNumber = acct[1];
  const client = text.match(/\b(?:for|client)\s+([A-Z][A-Za-z0-9 &.'-]{2,60}?)(?:,|\s+for\s+|\.|$)/);
  if (client) out.extractedClientName = client[1].trim();
  return out;
}

/** True for a fresh recon request OR a "change the layout of Bank Recon…" follow-up. */
export function isBankReconPrompt(message: string): boolean {
  return detectBankReconIntent(message) !== null || isBankReconRelayoutPrompt(message);
}

export function isBankReconRelayoutPrompt(message: string): boolean {
  const text = message.trim();
  if (!text || GST_MARKERS.test(text)) return false;
  return /\bbank\s+recon/i.test(text) && RELAYOUT_VERB.test(text) && RELAYOUT_WORDS.test(text);
}

/** Casual unless the user quoted an account number or asked for the full BRS report. */
export function isCasualBankRecon(message: string, intent: BankReconIntent): boolean {
  if (intent.extractedAccountNumber) return false;
  return !/\b(full\s+(?:recon|report|brs)|brs\s+report|detailed\s+report|account\s+(?:no|number))\b/i.test(message);
}

const BANK_ONLY =
  /\b(only\s+in\s+(?:the\s+)?bank|in\s+(?:the\s+)?bank\s+(?:but\s+)?not\s+(?:in\s+)?(?:my\s+|the\s+)?(?:books|ledger)|not\s+(?:been\s+)?(?:booked|recorded|posted)|entries\s+to\s+post|bank[-\s]?only|bank[-\s]?flat)\b/i;
const BOOKS_ONLY =
  /\b(only\s+in\s+(?:the\s+)?(?:books|ledger)|in\s+(?:my\s+|the\s+)?(?:books|ledger)\s+(?:but\s+)?not\s+(?:in\s+)?(?:the\s+)?bank|not\s+(?:yet\s+)?(?:cleared|presented|in\s+(?:the\s+)?bank)|books[-\s]?only|books[-\s]?flat|single\s+table|one\s+table|flat\s+table)\b/i;
const MATCHED = /\b(matched|audit\s+trail|what\s+matched)\b/i;

export function detectBankReconLayout(message: string): BankReconLayout {
  const text = message.trim();
  if (MATCHED.test(text) && !/\bnot\s+matched\b/i.test(text)) return 'matched';
  if (BANK_ONLY.test(text)) return 'bank_flat';
  if (BOOKS_ONLY.test(text)) return 'books_flat';
  return 'categorized';
}

/** "cheques not cleared / outstanding / deposits in transit" — only the timing items, not every books-only row. */
export function wantsTimingItemsOnly(message: string): boolean {
  return /\b(not\s+(?:yet\s+)?(?:cleared|presented|credited)|uncleared|outstanding|in\s+transit|stale)\b/i.test(message);
}

/** Layout asked for in a follow-up; falls back to books_flat when it only says "single table". */
export function detectBankReconRelayoutTarget(message: string): BankReconLayout {
  const text = message.trim();
  if (/\b(categorized|sections?|grouped?|by\s+reason)\b/i.test(text) && !/\b(single|one|flat)\s+table\b/i.test(text)) {
    return 'categorized';
  }
  if (MATCHED.test(text)) return 'matched';
  if (/\bbank[-\s]?(only|flat)\b/i.test(text) || BANK_ONLY.test(text)) return 'bank_flat';
  return 'books_flat';
}
