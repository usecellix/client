/**
 * Discover Purchase Register / GSTR sheets in the open workbook by
 * name hints + header signatures (not filename alone).
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
  portal: SheetCandidate | null;
  ims: SheetCandidate | null;
  /** Ambiguous candidates (same role, multiple sheets) */
  ambiguous: Array<{ role: DiscoveredSheetRole; sheets: string[] }>;
  missing: Array<'PURCHASE_REGISTER' | 'GSTR2B' | 'GSTR2A' | 'GSTR1' | 'IMS' | 'SALES_REGISTER'>;
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
      if (has('receiver') || has('b2b')) score += 2;
      break;
    }
    case 'PURCHASE_REGISTER': {
      if (name.includes('purchase') || name.includes(' pr ') || name.startsWith('pr')) score += 4;
      if (has('supplier gstin', 'party name', 'voucher', 'bill no')) score += 3;
      if (has('gstin') && has('invoice') && (has('taxable') || has('cgst') || has('igst')))
        score += 3;
      if (has('narration') || has('ledger')) score += 1;
      // Prefer not scoring GSTR portal dumps as PR
      if (name.includes('gstr') || name.includes('2b') || name.includes('2a')) score -= 6;
      break;
    }
    case 'SALES_REGISTER': {
      if (name.includes('sales') || name.includes('outward')) score += 4;
      if (has('receiver gstin', 'party name') && has('invoice')) score += 3;
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

/**
 * Infer columns for recon payload from header labels.
 */
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
  const gstin = find('gstin of supplier', 'supplier gstin', 'gstin', 'gst no');
  const invoiceNo = find(
    'invoice number',
    'invoice no',
    'inv no',
    'bill no',
    'voucher no',
    'document number',
  );
  const invoiceDate = find('invoice date', 'bill date', 'voucher date', 'date');
  const taxableAmt = find('taxable value', 'taxable amount', 'taxable amt', 'assessable value');
  const taxAmount = find('tax amount', 'total tax', 'total gst');
  const igst = find('igst');
  const cgst = find('cgst');
  const sgst = find('sgst', 'utgst');
  const narration = find('narration', 'description', 'particulars');
  const documentType = find('document type', 'doc type', 'invoice type', 'voucher type');
  const imsAction = find('ims action', 'ims status', 'recipient action', 'action');

  if (gstin !== undefined) mapping.gstin = gstin;
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

  const imsPick = pickBest(sheets, 'IMS', 4);

  if (!prPick.best && prPick.ties.length === 0) {
    missing.push(booksRole === 'SALES_REGISTER' ? 'SALES_REGISTER' : 'PURCHASE_REGISTER');
  }
  if (!portalPick.best && portalPick.ties.length === 0) {
    if (portalRole === 'GSTR2B') missing.push('GSTR2B');
    else if (portalRole === 'GSTR2A') missing.push('GSTR2A');
    else if (portalRole === 'GSTR1') missing.push('GSTR1');
    else missing.push('IMS');
  }

  return {
    purchaseRegister: prPick.best,
    portal: portalPick.best,
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
  if (missing.includes('GSTR2B')) {
    parts.push(
      "I couldn't find a **GSTR-2B** sheet in this workbook. Download GSTR-2B from the GST portal, paste/export it into a sheet in this file (headers should include supplier GSTIN and invoice columns), then ask me again.",
    );
  }
  if (missing.includes('GSTR2A')) {
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
