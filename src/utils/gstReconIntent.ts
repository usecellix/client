/**
 * Detect GST reconciliation intents from natural-language chat prompts.
 * Supports Purchase vs GSTR-2B/2A and Sales vs GSTR-1, plus context extraction.
 */

export type GstReconIntentType =
  | 'PR_VS_GSTR2B'
  | 'PR_VS_GSTR2A'
  | 'IMS_VS_PR'
  | 'SALES_VS_GSTR1'
  | 'GSTR3B_VS_GSTR2B';

export interface GstReconIntent {
  type: GstReconIntentType;
  /** True when phrasing is clearly GST recon / ITC match related */
  matched: boolean;
  extractedGstin?: string;
  extractedPeriod?: string;
  extractedFinancialYear?: string;
  extractedClientName?: string;
}

const RECON_VERB =
  /\b(reconcil(?:e|iation|ing)?|match(?:ing)?|compare|cross[-\s]?check|set[-\s]?off)\b/i;

const GST_MARKERS =
  /\b(gst|gstr|gstr[-\s]?2b|gstr[-\s]?2a|gstr[-\s]?1|2b|2a|itc|purchase\s+register|sales\s+register|ims|input\s+tax)\b/i;

const GSTIN_RE =
  /\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z])\b/i;

/**
 * Extract run-context fields from free text.
 */
export function extractGstReconContext(message: string): {
  extractedGstin?: string;
  extractedPeriod?: string;
  extractedFinancialYear?: string;
  extractedClientName?: string;
} {
  const text = message.trim();
  const out: {
    extractedGstin?: string;
    extractedPeriod?: string;
    extractedFinancialYear?: string;
    extractedClientName?: string;
  } = {};

  const gstinMatch = text.match(GSTIN_RE);
  if (gstinMatch) out.extractedGstin = gstinMatch[1].toUpperCase();

  const fy =
    text.match(/\b(?:FY|F\.?Y\.?)\s*(\d{2,4})\s*[-–/]\s*(\d{2,4})\b/i) ||
    text.match(/\b(20\d{2})\s*[-–/]\s*(\d{2,4})\b/);
  if (fy) {
    out.extractedFinancialYear = `${fy[1]}-${fy[2]}`;
  }

  const monthYear =
    text.match(
      /\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\s+(\d{4}|\d{2})\b/i,
    ) ||
    text.match(/\b(20\d{2})[-/](0?[1-9]|1[0-2])\b/) ||
    text.match(/\bQ([1-4])\s*(?:FY|of)?\s*(20\d{2}|\d{2})?\b/i);
  if (monthYear) {
    out.extractedPeriod = monthYear[0];
  }

  const client =
    text.match(
      /\b(?:for|client)\s+([A-Z][A-Za-z0-9 &.'-]{2,60}?)(?:,|\s+GSTIN|\s+gstin|\s+for\s+|\.|$)/,
    ) || text.match(/\b([A-Z][A-Za-z0-9 &.'-]{2,40}?)\s*,\s*GSTIN\b/i);
  if (client) {
    out.extractedClientName = client[1].trim().replace(/[,.]$/, '');
  }

  return out;
}

/**
 * Returns intent when the user is asking for GST-oriented reconciliation.
 */
export function detectGstReconIntent(message: string): GstReconIntent | null {
  const text = message.trim();
  if (!text) return null;

  const hasReconVerb = RECON_VERB.test(text);
  const hasGst = GST_MARKERS.test(text);

  const gstReconPhrase =
    /\bgst\s+recon(?:ciliation)?\b/i.test(text) ||
    /\brecon(?:cile|ciliation)?\s+(?:my\s+)?(?:gst|gstr|itc|purchases?|sales)\b/i.test(text) ||
    /\bmatch(?:\s+my)?\s+(?:purchases?|purchase\s+register|sales).*(?:gstr|2b|2a|1)\b/i.test(
      text,
    ) ||
    /\b(?:gstr|2b|2a).*(?:match|reconcil)/i.test(text) ||
    /\bpurchase\s+register.*(?:gstr|2b|2a)\b/i.test(text) ||
    /\bsales\s+register.*(?:gstr|gstr[-\s]?1)\b/i.test(text);

  if (!hasGst && !gstReconPhrase) return null;
  if (!hasReconVerb && !gstReconPhrase) return null;

  const ctx = extractGstReconContext(text);
  const base = { matched: true as const, ...ctx };

  if (/\b(3b|gstr[-\s]?3b)\b/i.test(text) && /\b(2b|gstr[-\s]?2b)\b/i.test(text)) {
    return { ...base, type: 'GSTR3B_VS_GSTR2B' };
  }
  if (/\bims\b/i.test(text)) {
    return { ...base, type: 'IMS_VS_PR' };
  }
  if (
    (/\b(sales|outward|gstr[-\s]?1)\b/i.test(text) && !/\b2b\b/i.test(text) && !/\b2a\b/i.test(text)) ||
    /\bsales\s+register\b/i.test(text)
  ) {
    return { ...base, type: 'SALES_VS_GSTR1' };
  }
  if (/\b(2a|gstr[-\s]?2a)\b/i.test(text)) {
    return { ...base, type: 'PR_VS_GSTR2A' };
  }

  return { ...base, type: 'PR_VS_GSTR2B' };
}

export function isGstReconPrompt(message: string): boolean {
  return detectGstReconIntent(message) !== null;
}

/**
 * Casual chat: "Reconcile purchase/sales register" — skip GSTIN/period prompts
 * and write only books rows missing on the portal.
 * Full GSTIN-gated recon still runs when a GSTIN is present or the prompt
 * is generic GST recon without a books-register phrase.
 */
export function isCasualMissedRowsRecon(
  message: string,
  intent: GstReconIntent,
): boolean {
  if (intent.type === 'IMS_VS_PR' || intent.type === 'GSTR3B_VS_GSTR2B') {
    return false;
  }
  if (intent.extractedGstin) return false;
  if (/\b(full\s+recon|recon(?:ciliation)?\s+report|itc\s+report)\b/i.test(message)) {
    return false;
  }
  return /\b(purchase\s+register|sales\s+register)\b/i.test(message);
}

/** Detect affirmative GSTIN confirmation replies. */
export function parseGstinConfirmation(message: string): string | null {
  const m = message.match(GSTIN_RE);
  if (m) return m[1].toUpperCase();
  return null;
}

/**
 * Output layout for the casual "missed rows" sheet.
 * - `categorized` (default): current multi-section layout, grouped by mismatch reason.
 * - `books_flat`: single flat table — books rows NOT found on the portal (any reason).
 * - `portal_flat`: the reverse — portal rows NOT found in the books.
 */
export type GstReconLayout = 'categorized' | 'books_flat' | 'portal_flat';

const BOOKS_FLAT_EXPLICIT = /\bjust\s+show\s+(?:the\s+|me\s+)?(?:the\s+)?missed\s+rows\b/i;
const SIMPLE_LAYOUT = /\bsimple\s+layout\b/i;

const BOOKS_SIDE = '(?:my\\s+)?(?:books|purchase\\s+register|sales\\s+register|register)';
const PORTAL_SIDE = '(?:the\\s+)?(?:portal|gstr[-\\s]?2a|gstr[-\\s]?2b|gstr[-\\s]?1)';

/**
 * "rows in <A> ... not in <B>" — whichever side is named first is the source set being
 * checked for existence; the side after "not in" is where it's absent. Getting this
 * backwards would silently swap the two comparison directions, so the mentioned ORDER
 * (not just which words appear) is what decides the layout.
 */
const BOOKS_THEN_PORTAL_MISSING = new RegExp(
  `\\b(?:rows?|invoices?)\\s+in\\s+${BOOKS_SIDE}\\b[^.?!]{0,40}?\\bnot\\s+in\\b[^.?!]{0,25}?\\b${PORTAL_SIDE}\\b`,
  'i',
);
const PORTAL_THEN_BOOKS_MISSING = new RegExp(
  `\\b(?:rows?|invoices?)\\s+in\\s+${PORTAL_SIDE}\\b[^.?!]{0,40}?\\bnot\\s+in\\b[^.?!]{0,25}?\\b${BOOKS_SIDE}\\b`,
  'i',
);
/** "invoices the portal has that I haven't recorded" — portal-side phrasing without "in ... not in". */
const PORTAL_HAS_NOT_RECORDED =
  /\b(?:the\s+)?portal\s+has\b[^.?!]{0,50}?\b(?:i\s+(?:haven'?t|have\s+not)\s+recorded|not\s+(?:in\s+)?(?:my\s+)?books)\b/i;

/**
 * Returns null (never a layout guess) for anything ambiguous — e.g. "what's missing in
 * GSTR-2B" names only one side and no direction, and must stay on the categorized
 * default rather than being misread as either flat layout.
 */
export function detectGstReconLayout(message: string): GstReconLayout {
  const text = message.trim();
  if (!text) return 'categorized';

  if (BOOKS_FLAT_EXPLICIT.test(text) || SIMPLE_LAYOUT.test(text)) return 'books_flat';
  if (PORTAL_THEN_BOOKS_MISSING.test(text) || PORTAL_HAS_NOT_RECORDED.test(text)) {
    return 'portal_flat';
  }
  if (BOOKS_THEN_PORTAL_MISSING.test(text)) return 'books_flat';

  return 'categorized';
}
