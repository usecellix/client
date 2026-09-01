/**
 * Detect GST reconciliation intents from natural-language chat prompts.
 * MVP focus: Purchase Register vs GSTR-2B (Type 01).
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
}

const RECON_VERB =
  /\b(reconcil(?:e|iation|ing)?|match(?:ing)?|compare|cross[-\s]?check|set[-\s]?off)\b/i;

const GST_MARKERS =
  /\b(gst|gstr|gstr[-\s]?2b|gstr[-\s]?2a|gstr[-\s]?1|2b|2a|itc|purchase\s+register|ims|input\s+tax)\b/i;

/**
 * Returns intent when the user is asking for GST-oriented reconciliation.
 */
export function detectGstReconIntent(message: string): GstReconIntent | null {
  const text = message.trim();
  if (!text) return null;

  const hasReconVerb = RECON_VERB.test(text);
  const hasGst = GST_MARKERS.test(text);

  // Explicit combined phrases without strict verb (e.g. "GST recon for April")
  const gstReconPhrase =
    /\bgst\s+recon(?:ciliation)?\b/i.test(text) ||
    /\brecon(?:cile|ciliation)?\s+(?:my\s+)?(?:gst|gstr|itc|purchases?)\b/i.test(text) ||
    /\bmatch(?:\s+my)?\s+(?:purchases?|purchase\s+register|sales).*(?:gstr|2b|2a)\b/i.test(text) ||
    /\b(?:gstr|2b|2a).*(?:match|reconcil)/i.test(text) ||
    /\bpurchase\s+register.*(?:gstr|2b|2a)\b/i.test(text);

  if (!hasGst && !gstReconPhrase) return null;
  if (!hasReconVerb && !gstReconPhrase) return null;

  if (/\b(3b|gstr[-\s]?3b)\b/i.test(text) && /\b(2b|gstr[-\s]?2b)\b/i.test(text)) {
    return { matched: true, type: 'GSTR3B_VS_GSTR2B' };
  }
  if (/\bims\b/i.test(text)) {
    return { matched: true, type: 'IMS_VS_PR' };
  }
  if (/\b(sales|outward|gstr[-\s]?1)\b/i.test(text) && !/\b2b\b/i.test(text)) {
    return { matched: true, type: 'SALES_VS_GSTR1' };
  }
  if (/\b(2a|gstr[-\s]?2a)\b/i.test(text)) {
    return { matched: true, type: 'PR_VS_GSTR2A' };
  }

  // Default MVP: purchase vs GSTR-2B
  return { matched: true, type: 'PR_VS_GSTR2B' };
}

export function isGstReconPrompt(message: string): boolean {
  return detectGstReconIntent(message) !== null;
}
