import { describe, expect, it } from 'vitest';
import { gstinCheckChar, isValidGstin } from './indiaValidators';

// Same fixtures as T_GSTIN / Server india-validators.util.spec.ts
const VALID = [
  '27AABCA1234F1Z9', '29AAACM5678K1ZV', '07BBBPS4321Q1ZA', '33CCCHD9876L1ZN',
  '24DDDFR1111M1ZA', '19EEETP2222N1ZY', '36FFFCT3333P1ZJ', '09GGGAS4444R1Z7',
];
const INVALID = [
  '27AABCA1234F1Z7', '29AAACA5678K1ZV', '07BBBPS4371Q1ZA', '37CCCHD9876L1ZN',
  '24DDDFR1111M', '50HHHPT5555S1ZC', '', 'not a gstin',
];

describe('client GSTIN check (mirror of the server B10 module)', () => {
  it.each(VALID)('valid: %s', (g) => expect(isValidGstin(g)).toBe(true));
  it.each(INVALID)('invalid: %j', (g) => expect(isValidGstin(g)).toBe(false));
  it('ignores case and stray spaces, never alters the input', () => {
    expect(isValidGstin(' 27aabca1234f1z9 ')).toBe(true);
  });
  it('check char of the first 14 matches the 15th', () => {
    for (const g of VALID) expect(gstinCheckChar(g.slice(0, 14))).toBe(g[14]);
  });
});
