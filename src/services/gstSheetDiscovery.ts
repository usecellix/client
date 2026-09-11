/**
 * Discover Purchase Register / Sales Register / GSTR sheets in the open workbook by
 * name hints + header signatures (aligned with Server sheet-detector HEADER_ALIASES).
 */

export type DiscoveredSheetRole =
  | 'PURCHASE_REGISTER'
  | 'SALES_REGISTER'
  | 'GSTR2B'
  | 'GSTR2A'
  | 'GSTR1'
  | 'IMS'
  | 'UNKNOWN';

export interface SheetCandidate {
  name: string;
  role: DiscoveredSheetRole;
  score: number;
  headers: string[];
}

export interface GstSheetDiscoveryResult {
  purchaseRegister: SheetCandidate | null;
  /** Alias — books-side register (PR or Sales). */
  booksRegister: SheetCandidate | null;
  portal: SheetCandidate | null;
  gstr2b: SheetCandidate | null;
  gstr2a: SheetCandidate | null;
  ims: SheetCandidate | null;
  ambiguous: Array<{ role: DiscoveredSheetRole; sheets: string[] }>;
  missing: Array<
    'PURCHASE_REGISTER' | 'GSTR2B' | 'GSTR2A' | 'GSTR1' | 'IMS' | 'SALES_REGISTER'
  >;
}

function norm(h: unknown): string {
  return String(h ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-./]+/g, ' ');
}

function scoreRole(
  sheetName: string,
  headers: string[],
  role: DiscoveredSheetRole,
): number {
  const name = norm(sheetName);
  const joined = headers.map(norm).join(' | ');
  let score = 0;

  const has = (...keys: string[]) => keys.some((k) => joined.includes(k) || name.includes(k));

  switch (role) {
    case 'GSTR2B': {
      if (name.includes('2b') || name.includes('gstr2b') || name.includes('gstr 2b')) score += 5;
      if (has('gstin of supplier', 'itc available', 'document type')) score += 3;
      if (has('taxable value') && has('gstin')) score += 2;
      if (has('cgst') || has('sgst') || has('igst')) score += 1;
      break;
    }
    case 'GSTR2A': {
      if (name.includes('2a') || name.includes('gstr2a')) score += 5;
      if (has('tds') || has('tcs')) score += 2;
      if (has('gstin') && has('invoice')) score += 1;
      break;
    }
    case 'GSTR1': {
      if (name.includes('gstr1') || name.includes('gstr 1') || /\bgstr-?1\b/.test(name))
        score += 5;
      if (has('receiver', 'recipient gstin', 'gstin of recipient', 'b2b', 'irn')) score += 3;
      if (has('invoice') && has('taxable')) score += 1;
      break;
    }
    case 'PURCHASE_REGISTER': {
      if (name.includes('purchase') || name.includes(' pr ') || name.startsWith('pr')) score += 4;
      if (has('supplier gstin', 'gstin of supplier', 'party name', 'voucher', 'bill no'))
        score += 3;
      if (has('gstin') && has('invoice') && (has('taxable') || has('cgst') || has('igst')))
        score += 3;
      if (has('narration') || has('ledger')) score += 1;
      if (has('recipient gstin', 'gstin of recipient') && !has('supplier gstin')) score -= 3;
      if (name.includes('gstr') || name.includes('2b') || name.includes('2a')) score -= 6;
      break;
    }
    case 'SALES_REGISTER': {
      if (name.includes('sales') || name.includes('outward')) score += 4;
      if (has('recipient gstin', 'receiver gstin', 'gstin of recipient', 'customer gstin'))
        score += 4;
      if (has('supply type', 'supply category', 'b2b', 'b2c')) score += 2;
      if (has('invoice') && (has('taxable') || has('cgst'))) score += 2;
      if (has('supplier gstin') && !has('recipient gstin')) score -= 2;
      if (name.includes('gstr')) score -= 4;
      break;
    }
    case 'IMS': {
      if (name.includes('ims')) score += 5;
      if (has('ims action', 'ims status', 'accept', 'reject', 'pending')) score += 3;
      break;
    }
    default:
      break;
  }
  return score;
}

function pickBest(
  sheets: Array<{ name: string; headers: string[] }>,
  role: DiscoveredSheetRole,
  minScore = 3,
): { best: SheetCandidate | null; ties: string[] } {
  const scored: SheetCandidate[] = sheets
    .map((s) => ({
      name: s.name,
      role,
      score: scoreRole(s.name, s.headers, role),
      headers: s.headers,
    }))
    .filter((s) => s.score >= minScore)
    .sort((a, b) => b.score - a.score);

  if (!scored.length) return { best: null, ties: [] };

  const top = scored[0].score;
  const ties = scored.filter((s) => s.score === top).map((s) => s.name);
  if (ties.length > 1) {
    return { best: null, ties };
  }
  return { best: scored[0], ties: [] };
}

export function inferColumnMapping(headers: string[]): Record<string, number> {
  const norms = headers.map(norm);
  const find = (...aliases: string[]) => {
    for (const a of aliases) {
      const i = norms.findIndex((h) => h === a || h.includes(a));
      if (i >= 0) return i;
    }
    return undefined;
  };

  const mapping: Record<string, number> = {};
  const gstin = find(
    'gstin of supplier',
    'supplier gstin',
    'gstin of recipient',
    'recipient gstin',
    'receiver gstin',
    'customer gstin',
    'gstin/uin of recipient',
    'gstin',
    'gst no',
  );
  const clientGstin = find(
    'client gstin',
    'our gstin',
    'company gstin',
    'gstin of registered person',
  );
  const invoiceNo = find(
    'invoice number',
    'invoice no',
    'inv no',
    'bill no',
    'voucher no',
    'document number',
  );
  const invoiceDate = find('invoice date', 'bill date', 'voucher date', 'document date', 'date');
  const taxableAmt = find('taxable value', 'taxable amount', 'taxable amt', 'assessable value');
  const taxAmount = find('tax amount', 'total tax', 'total gst');
  const igst = find('igst');
  const cgst = find('cgst');
  const sgst = find('sgst', 'utgst');
  const narration = find('narration', 'description', 'particulars');
  const documentType = find('document type', 'doc type', 'invoice type', 'voucher type');
  const imsAction = find('ims action', 'ims status', 'recipient action', 'action');
  const irn = find('irn', 'e-invoice irn', 'invoice reference number');
  const supplyCategory = find('supply category', 'supply type', 'b2b/b2c');
  const placeOfSupply = find('place of supply', 'pos');

  if (gstin !== undefined) mapping.gstin = gstin;
  if (clientGstin !== undefined) mapping.clientGstin = clientGstin;
  if (invoiceNo !== undefined) mapping.invoiceNo = invoiceNo;
  if (invoiceDate !== undefined) mapping.invoiceDate = invoiceDate;
  if (taxableAmt !== undefined) mapping.taxableAmt = taxableAmt;
  if (taxAmount !== undefined) mapping.taxAmount = taxAmount;
  if (igst !== undefined) mapping.igst = igst;
  if (cgst !== undefined) mapping.cgst = cgst;
  if (sgst !== undefined) mapping.sgst = sgst;
  if (narration !== undefined) mapping.narration = narration;
  if (documentType !== undefined) mapping.documentType = documentType;
  if (imsAction !== undefined) mapping.imsAction = imsAction;
  if (irn !== undefined) mapping.irn = irn;
  if (supplyCategory !== undefined) mapping.supplyCategory = supplyCategory;
  if (placeOfSupply !== undefined) mapping.placeOfSupply = placeOfSupply;
  return mapping;
}

export function discoverGstSheets(
  sheets: Array<{ name: string; headers: string[] }>,
  preferredPortal: 'GSTR2B' | 'GSTR2A' | 'GSTR1' | 'IMS' = 'GSTR2B',
): GstSheetDiscoveryResult {
  const ambiguous: GstSheetDiscoveryResult['ambiguous'] = [];
  const missing: GstSheetDiscoveryResult['missing'] = [];

  const booksRole: DiscoveredSheetRole =
    preferredPortal === 'GSTR1' ? 'SALES_REGISTER' : 'PURCHASE_REGISTER';

  const prPick = pickBest(sheets, booksRole);
  if (prPick.ties.length > 1) {
    ambiguous.push({ role: booksRole, sheets: prPick.ties });
  }

  const portalRole: DiscoveredSheetRole =
    preferredPortal === 'GSTR2A'
      ? 'GSTR2A'
      : preferredPortal === 'GSTR1'
        ? 'GSTR1'
        : preferredPortal === 'IMS'
          ? 'IMS'
          : 'GSTR2B';

  const portalPick = pickBest(sheets, portalRole);
  if (portalPick.ties.length > 1) {
    ambiguous.push({ role: portalRole, sheets: portalPick.ties });
  }

  const emptyPick: { best: SheetCandidate | null; ties: string[] } = { best: null, ties: [] };
  const imsPick = pickBest(sheets, 'IMS', 4);
  const gstr2bPick =
    booksRole === 'PURCHASE_REGISTER' && portalRole !== 'IMS'
      ? pickBest(sheets, 'GSTR2B')
      : emptyPick;
  const gstr2aPick =
    booksRole === 'PURCHASE_REGISTER' && portalRole !== 'IMS'
      ? pickBest(sheets, 'GSTR2A')
      : emptyPick;
  if (gstr2bPick.ties.length > 1 && portalRole !== 'GSTR2B') {
    ambiguous.push({ role: 'GSTR2B', sheets: gstr2bPick.ties });
  }
  if (gstr2aPick.ties.length > 1 && portalRole !== 'GSTR2A') {
    ambiguous.push({ role: 'GSTR2A', sheets: gstr2aPick.ties });
  }

  let gstr2bBest = gstr2bPick.best;
  let gstr2aBest = gstr2aPick.best;
  const nameHint = (n: string) => norm(n);
  if (
    gstr2bBest &&
    (nameHint(gstr2bBest.name).includes('2a') || nameHint(gstr2bBest.name).includes('gstr2a')) &&
    !nameHint(gstr2bBest.name).includes('2b')
  ) {
    gstr2bBest = null;
    if (!gstr2aBest) gstr2aBest = gstr2bPick.best;
  }

  if (!prPick.best && prPick.ties.length === 0) {
    missing.push(booksRole === 'SALES_REGISTER' ? 'SALES_REGISTER' : 'PURCHASE_REGISTER');
  }
  if (portalRole === 'IMS') {
    if (!portalPick.best && portalPick.ties.length === 0) missing.push('IMS');
  } else if (booksRole === 'PURCHASE_REGISTER') {
    if (!gstr2bBest && !gstr2aBest) {
      missing.push('GSTR2B');
      missing.push('GSTR2A');
    }
  } else if (!portalPick.best && portalPick.ties.length === 0) {
    missing.push('GSTR1');
  }

  const purchasePortal = portalPick.best ?? gstr2bBest ?? gstr2aBest ?? null;

  return {
    purchaseRegister: prPick.best,
    booksRegister: prPick.best,
    portal: booksRole === 'PURCHASE_REGISTER' ? purchasePortal : portalPick.best,
    gstr2b: gstr2bBest,
    gstr2a: gstr2aBest,
    ims: imsPick.best,
    ambiguous,
    missing,
  };
}

export function buildMissingSheetMessage(
  missing: GstSheetDiscoveryResult['missing'],
  intentLabel: string,
): string {
  const parts: string[] = [];
  if (missing.includes('GSTR2B') && missing.includes('GSTR2A')) {
    parts.push(
      "I couldn't find a **GSTR-2B** or **GSTR-2A** sheet in this workbook. Download either (or both) from the GST portal, paste them into this file, then ask me again.",
    );
  } else if (missing.includes('GSTR2B')) {
    parts.push(
      "I couldn't find a **GSTR-2B** sheet in this workbook. Download GSTR-2B from the GST portal, paste/export it into a sheet in this file (headers should include supplier GSTIN and invoice columns), then ask me again.",
    );
  } else if (missing.includes('GSTR2A')) {
    parts.push(
      "I couldn't find a **GSTR-2A** sheet. Add your GSTR-2A download as a sheet, then ask again.",
    );
  }
  if (missing.includes('GSTR1')) {
    parts.push(
      "I couldn't find a **GSTR-1** sheet. Add the GSTR-1 export as a sheet, then ask again.",
    );
  }
  if (missing.includes('PURCHASE_REGISTER')) {
    parts.push(
      "I couldn't find a **Purchase Register** (books) sheet with GSTIN / invoice / taxable columns. Ensure your purchase data is on a sheet in this workbook.",
    );
  }
  if (missing.includes('SALES_REGISTER')) {
    parts.push(
      "I couldn't find a **Sales Register** sheet with invoice and tax columns. Add it to this workbook and try again.",
    );
  }
  if (missing.includes('IMS')) {
    parts.push(
      "I couldn't find an **IMS** sheet. Download IMS data from the GST portal into Excel and ask again.",
    );
  }
  if (!parts.length) {
    return `I couldn't locate the sheets needed for ${intentLabel}.`;
  }
  return parts.join('\n\n');
}

export function buildAmbiguousSheetMessage(
  ambiguous: GstSheetDiscoveryResult['ambiguous'],
): string {
  return ambiguous
    .map(
      (a) =>
        `I found multiple possible **${a.role.replace(/_/g, ' ')}** sheets: ${a.sheets
          .map((s) => `"${s}"`)
          .join(', ')}. ` +
        'Please tell me which one to use (e.g. “use Purchases Apr24”).',
    )
    .join('\n\n');
}
