import { afterEach, describe, expect, it } from 'vitest';
import { detectRecipeIntent, headerRowOf, listBankRecipes, matchRecipe, rememberBankReconRun, saveBankRecipe } from './bankRecipes';
import { isBankReconPrompt } from '@/utils/bankReconIntent';

const store: Record<string, string> = {};
(globalThis as unknown as { localStorage: Storage }).localStorage = {
  getItem: (k: string) => store[k] ?? null,
  setItem: (k: string, v: string) => {
    store[k] = v;
  },
} as Storage;

const BANK = [['Account: 1234'], ['Date', 'Narration', 'Withdrawal Amt.', 'Deposit Amt.', 'Closing Balance'], ['01/09/26', 'NEFT X', 10, '', 90]];
const BOOKS = [['Date', 'Particulars', 'Vch Type', 'Debit', 'Credit'], ['01-09-2026', 'X', 'Payment', '', 10]];

describe('P4.5 bank recon recipes', () => {
  afterEach(() => Object.keys(store).forEach((k) => delete store[k]));

  it('finds the header row below account details', () => {
    expect(headerRowOf(BANK)).toEqual(['Date', 'Narration', 'Withdrawal Amt.', 'Deposit Amt.', 'Closing Balance']);
  });

  it('saves header names only — never values', () => {
    rememberBankReconRun(BANK, BOOKS, 'categorized');
    const r = saveBankRecipe('HDFC monthly')!;
    expect(r.bankHeaders).toContain('Narration');
    expect(store['cellix.bankRecipes.v1']).not.toMatch(/NEFT X|Payment|Account: 1234/);
    expect(listBankRecipes()).toHaveLength(1);
  });

  it('replay matches by header NAME even when columns are reordered', () => {
    rememberBankReconRun(BANK, BOOKS);
    const r = saveBankRecipe('HDFC monthly')!;
    const reordered = [
      { name: 'Oct Statement', rows: [['Closing Balance', 'Deposit Amt.', 'Narration', 'Withdrawal Amt.', 'Date']] },
      { name: 'Oct Ledger', rows: [['Credit', 'Debit', 'Vch Type', 'Particulars', 'Date']] },
    ];
    expect(matchRecipe(r, reordered)).toEqual({ bank: 'Oct Statement', books: 'Oct Ledger', missing: [] });
  });

  it('a missing header stops with a clear message', () => {
    rememberBankReconRun(BANK, BOOKS);
    const r = saveBankRecipe('HDFC monthly')!;
    const tops = [
      { name: 'Oct Statement', rows: [['Date', 'Description', 'Withdrawal Amt.', 'Deposit Amt.', 'Closing Balance']] },
      { name: 'Oct Ledger', rows: [['Date', 'Particulars', 'Vch Type', 'Debit', 'Credit']] },
    ];
    const m = matchRecipe(r, tops);
    expect(m.bank).toBeNull();
    expect(m.missing).toEqual(['bank statement column "Narration"']);
  });

  it('routes save / replay prompts to the bank recon chat', () => {
    expect(detectRecipeIntent('Save this as a recipe named HDFC monthly')).toEqual({ kind: 'save', name: 'HDFC monthly' });
    expect(detectRecipeIntent('Replay recipe HDFC monthly for October 2026')?.kind).toBe('replay');
    expect(isBankReconPrompt('Replay the recipe')).toBe(true);
    expect(isBankReconPrompt('Write a recipe for pasta')).toBe(false);
  });
});
