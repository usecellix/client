/**
 * P4.5 / A18.1 — bank-recon "recipe replay" (local stand-in for the colleague's B8 recipes, C3.6).
 *
 * A recipe is the SHAPE of a reconciliation that worked: which header names the bank
 * statement and the books ledger had, and the layout asked for. It holds header names only —
 * never amounts, narrations or any cell value — and lives in this browser's storage.
 * Replay finds next month's sheets by header NAME (column order does not matter); a missing
 * header stops with a clear message instead of guessing.
 */
import type { SheetTop } from './bankSheetDiscovery';

export interface BankRecipe {
  name: string;
  bankHeaders: string[];
  booksHeaders: string[];
  layout?: string;
  savedAt: string;
}

const KEY = 'cellix.bankRecipes.v1';
let lastRun: Omit<BankRecipe, 'name' | 'savedAt'> | null = null;

const norm = (h: unknown) => String(h ?? '').trim().toLowerCase().replace(/\s+/g, ' ');

/** The header row: the first row (within the top rows) that has a "date"-like cell and ≥ 3 filled cells. */
export function headerRowOf(rows: unknown[][]): string[] {
  for (const row of rows.slice(0, 40)) {
    const cells = (row ?? []).map((c) => String(c ?? '').trim());
    if (cells.filter(Boolean).length >= 3 && cells.some((c) => /\bdate\b|\bdt\b/i.test(c))) return cells.filter(Boolean);
  }
  return [];
}

export function rememberBankReconRun(bankRows: unknown[][], booksRows: unknown[][], layout?: string): void {
  lastRun = { bankHeaders: headerRowOf(bankRows), booksHeaders: headerRowOf(booksRows), layout };
}

export function listBankRecipes(): BankRecipe[] {
  try {
    const raw = globalThis.localStorage?.getItem(KEY);
    const parsed = raw ? (JSON.parse(raw) as unknown) : [];
    return Array.isArray(parsed) ? (parsed as BankRecipe[]).filter((r) => r && typeof r.name === 'string') : [];
  } catch {
    return [];
  }
}

export function saveBankRecipe(name: string): BankRecipe | null {
  if (!lastRun || !lastRun.bankHeaders.length || !lastRun.booksHeaders.length) return null;
  const recipe: BankRecipe = { name: name.trim().slice(0, 60) || 'Bank recon', ...lastRun, savedAt: new Date().toISOString() };
  try {
    const rest = listBankRecipes().filter((r) => r.name.toLowerCase() !== recipe.name.toLowerCase());
    globalThis.localStorage?.setItem(KEY, JSON.stringify([...rest, recipe]));
    return recipe;
  } catch {
    return null;
  }
}

export interface RecipeMatch {
  bank: string | null;
  books: string | null;
  missing: string[];
}

/** Find the sheets whose header row contains EVERY header of the recipe (any order). */
export function matchRecipe(recipe: BankRecipe, tops: SheetTop[]): RecipeMatch {
  const score = (headers: string[], wanted: string[]) => {
    const have = new Set(headers.map(norm));
    return wanted.filter((w) => !have.has(norm(w)));
  };
  let bank: string | null = null;
  let books: string | null = null;
  let bestBankMissing: string[] = recipe.bankHeaders;
  let bestBooksMissing: string[] = recipe.booksHeaders;
  for (const t of tops) {
    const h = headerRowOf(t.rows);
    const mb = score(h, recipe.bankHeaders);
    const ml = score(h, recipe.booksHeaders);
    if (mb.length < bestBankMissing.length) {
      bestBankMissing = mb;
      if (!mb.length) bank = t.name;
    }
    if (ml.length < bestBooksMissing.length) {
      bestBooksMissing = ml;
      if (!ml.length) books = t.name;
    }
  }
  return {
    bank,
    books: books === bank ? null : books,
    missing: [
      ...(bank ? [] : bestBankMissing.map((m) => `bank statement column "${m}"`)),
      ...(books && books !== bank ? [] : bestBooksMissing.map((m) => `ledger column "${m}"`)),
    ],
  };
}

export function detectRecipeIntent(message: string): { kind: 'save' | 'replay'; name: string } | null {
  const save = /\bsave\b.*\bas\s+(?:a\s+)?recipe\b(?:\s+(?:named|called)\s+["“']?([^"”']+)["”']?)?/i.exec(message);
  if (save) return { kind: 'save', name: save[1]?.trim() || 'Bank recon' };
  const replay = /\b(replay|re-?run|run)\b.*\brecipe\b(?:\s+["“']?([^"”']+?)["”']?)?\s*$/i.exec(message);
  if (replay) return { kind: 'replay', name: replay[2]?.trim() ?? '' };
  return null;
}
