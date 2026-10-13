import { describe, expect, it } from 'vitest';
import { discoverBankSheets } from './bankSheetDiscovery';

const statement = {
  name: 'Sheet1',
  rows: [
    ['HDFC BANK LTD'],
    ['Account No : 50100123456789   IFSC: HDFC0000123'],
    [],
    ['Date', 'Narration', 'Chq./Ref.No.', 'Value Dt', 'Withdrawal Amt.', 'Deposit Amt.', 'Closing Balance'],
    ['02/04/25', 'NEFT-ABC', 'X1', '02/04/25', 100, '', 900],
  ],
};
const ledger = {
  name: 'Bank A/c',
  rows: [
    ['Date', 'Particulars', 'Vch Type', 'Vch No.', 'Debit', 'Credit'],
    ['02-04-2025', 'ABC', 'Payment', 1, '', 100],
  ],
};
const other = { name: 'Notes', rows: [['hello'], ['world']] };

describe('discoverBankSheets', () => {
  it('finds both sheets even when the statement header is buried under account details', () => {
    expect(discoverBankSheets([ledger, statement, other])).toMatchObject({ bank: 'Sheet1', books: 'Bank A/c' });
  });

  it('reports ambiguity instead of guessing', () => {
    const second = { ...statement, name: 'Sheet2' };
    const r = discoverBankSheets([ledger, statement, second]);
    expect(r.bank).toBeNull();
    expect(r.ambiguousBank.sort()).toEqual(['Sheet1', 'Sheet2']);
  });

  it('returns nulls when there is no bank data', () => {
    expect(discoverBankSheets([other])).toMatchObject({ bank: null, books: null });
  });
});

describe('discovery edge cases', () => {
  it('ignores its own output sheet and finds the prior BRS', async () => {
    const { findPriorBrsSheet } = await import('./bankSheetDiscovery');
    const out = { name: 'Bank Recon Summary', rows: [['Date', 'Narration', 'Receipt', 'Payment', 'Balance'], ['2026-09-02', 'x', 1, '', 5]] };
    const r = discoverBankSheets([ledger, statement, out]);
    expect(r).toMatchObject({ bank: 'Sheet1', books: 'Bank A/c' });
    expect(findPriorBrsSheet([{ name: 'Prior BRS (Aug-26)' }, { name: 'Bank Recon Summary' }, { name: 'Sheet1' }])).toBe('Prior BRS (Aug-26)');
    expect(findPriorBrsSheet([{ name: 'Bank Statement' }])).toBeNull();
  });
});
