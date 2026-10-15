/* global Excel */

/**
 * Finds the bank statement and the books bank ledger among the workbook's sheets.
 * Bank statements put account details above the real header row, so — unlike GST sheet
 * discovery — this inspects the top rows of each sheet, not just row 1.
 */

export interface SheetTop {
  name: string;
  rows: unknown[][];
}

export interface BankSheetDiscovery {
  bank: string | null;
  books: string | null;
  /** Several sheets look equally like the bank statement / the ledger. */
  ambiguousBank: string[];
  ambiguousBooks: string[];
  /** Every sheet that looks like a transaction table (own output and prior-BRS sheets excluded). */
  candidates: string[];
}

const HEADER_PATTERNS: Array<[string, RegExp]> = [
  ['valueDate', /value ?(date|dt)/],
  ['vch', /^(vch|voucher) ?(type|no|num|number)/],
  ['date', /^(txn |tran |trans |transaction |posting |post |voucher |vch |book |booking )?(date|dt)$/],
  ['moneyOut', /^(debit|withdrawal|withdrawals|dr|paid out|payment|payments)\b/],
  ['moneyIn', /^(credit|deposit|deposits|cr|paid in|receipt|receipts)\b/],
  ['amount', /^(txn |tran |transaction )?amount( in rs)?$/],
  ['balance', /balance/],
  ['narration', /narration|description|particulars|remarks|details/],
];

function normHeader(raw: unknown): string {
  return String(raw ?? '')
    .toLowerCase()
    .replace(/[._\-:()]/g, ' ')
    .replace(/\s*\/\s*/g, '/')
    .replace(/\s+/g, ' ')
    .trim();
}

interface SheetScore {
  name: string;
  hasHeader: boolean;
  /** Positive = bank statement, negative = books ledger. */
  score: number;
}

function scoreSheet(sheet: SheetTop): SheetScore {
  const name = sheet.name.toLowerCase();
  let score = 0;
  if (/statement|passbook|stmt/.test(name)) score += 1;
  if (/ledger|cash\s*book|bank\s*book|books|tally|zoho|busy/.test(name)) score -= 1;

  let best: { idx: number; kinds: Set<string>; cells: string[] } | null = null;
  sheet.rows.slice(0, 40).forEach((row, idx) => {
    const kinds = new Set<string>();
    const cells = (row ?? []).map(normHeader);
    for (const c of cells) {
      if (!c || c.length > 40) continue;
      for (const [kind, re] of HEADER_PATTERNS) {
        if (re.test(c)) {
          kinds.add(kind);
          break;
        }
      }
    }
    const ok = kinds.has('date') && (kinds.has('moneyOut') || kinds.has('moneyIn') || kinds.has('amount'));
    if (ok && (!best || kinds.size > best.kinds.size)) best = { idx, kinds, cells };
  });

  if (!best) return { name: sheet.name, hasHeader: false, score };
  const { kinds, cells, idx } = best as { idx: number; kinds: Set<string>; cells: string[] };
  if (kinds.has('vch')) score -= 2;
  if (kinds.has('valueDate')) score += 2;
  if (cells.some((c) => /^(withdrawal|deposit)/.test(c))) score += 2;
  if (cells.includes('particulars') && !kinds.has('balance')) score -= 1;
  if (kinds.has('balance') && !kinds.has('vch')) score += 1;
  const pre = sheet.rows
    .slice(0, idx)
    .map((r) => (r ?? []).join(' '))
    .join(' ');
  if (/\bifsc\b|account\s*(no|number)|a\/c\s*no|customer\s*id/i.test(pre)) score += 1;
  return { name: sheet.name, hasHeader: true, score };
}

/** Role of one specific sheet, for a sheet the user named with @[Sheet]. */
export function classifySheetRole(sheet: SheetTop): 'bank' | 'books' | 'unknown' {
  const s = scoreSheet(sheet);
  if (!s.hasHeader) return 'unknown';
  return s.score > 0 ? 'bank' : s.score < 0 ? 'books' : 'unknown';
}

/** Sheets this feature writes itself — never candidates for the statement or ledger on a re-run. */
const OUTPUT_SHEET_RE = /^bank recon\b/i;

/** Previous period's BRS ("Prior BRS (Aug-26)", "Previous BRS", "Last month BRS"). */
const PRIOR_BRS_RE = /\b(prior|previous|last|opening|carry)\b.*\b(brs|reconciliation)\b|\b(brs|reconciliation)\b.*\b(prior|previous|aug|carry)/i;

export function findPriorBrsSheet(sheets: Array<{ name: string }>): string | null {
  const hits = sheets.filter((s) => PRIOR_BRS_RE.test(s.name) && !OUTPUT_SHEET_RE.test(s.name));
  return hits.length === 1 ? hits[0].name : null;
}

export function discoverBankSheets(sheets: SheetTop[]): BankSheetDiscovery {
  const scored = sheets
    .filter((s) => !OUTPUT_SHEET_RE.test(s.name))
    .map(scoreSheet)
    .filter((s) => s.hasHeader);
  const pick = (cmp: (s: SheetScore) => boolean, rank: (s: SheetScore) => number) => {
    const c = scored.filter(cmp).sort((a, b) => rank(b) - rank(a));
    if (!c.length) return { best: null as string | null, ties: [] as string[] };
    const top = rank(c[0]);
    const ties = c.filter((s) => rank(s) === top).map((s) => s.name);
    return ties.length > 1 ? { best: null, ties } : { best: c[0].name, ties: [] };
  };
  const bank = pick((s) => s.score >= 2, (s) => s.score);
  const books = pick((s) => s.score <= -2, (s) => -s.score);

  // Exactly two usable sheets and one is unambiguous: the other one is the counterpart.
  if (!bank.best && !bank.ties.length && books.best) {
    const rest = scored.filter((s) => s.name !== books.best);
    if (rest.length === 1) bank.best = rest[0].name;
  }
  if (!books.best && !books.ties.length && bank.best) {
    const rest = scored.filter((s) => s.name !== bank.best);
    if (rest.length === 1) books.best = rest[0].name;
  }
  return {
    bank: bank.best,
    books: books.best,
    ambiguousBank: bank.ties,
    ambiguousBooks: books.ties,
    candidates: scored.map((s) => s.name),
  };
}

/** First `rowCount` rows of every sheet — enough to find a header buried under account details. */
export async function readSheetTops(rowCount = 40, maxSheets = 30): Promise<SheetTop[]> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();
    const out: SheetTop[] = [];
    for (const s of sheets.items.slice(0, maxSheets)) {
      try {
        const used = s.getUsedRange();
        used.load(['rowCount', 'columnCount']);
        await ctx.sync();
        if (!used.rowCount || !used.columnCount) {
          out.push({ name: s.name, rows: [] });
          continue;
        }
        const range = s.getRangeByIndexes(0, 0, Math.min(rowCount, used.rowCount), used.columnCount);
        range.load('values');
        await ctx.sync();
        out.push({ name: s.name, rows: (range.values as unknown[][]) ?? [] });
      } catch {
        out.push({ name: s.name, rows: [] });
      }
    }
    return out;
  });
}
