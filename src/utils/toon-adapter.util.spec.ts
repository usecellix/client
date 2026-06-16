import { describe, expect, it } from 'vitest';
import { estimateTokens, fromToon, toToon } from '@/utils/toon-adapter.util';

describe('toon-adapter', () => {
  it('falls back to JSON for very small row sets', () => {
    const rows = [
      { name: 'A', qty: 1 },
      { name: 'B', qty: 2 },
    ];
    const encoded = toToon(rows);
    expect(encoded).toBe(JSON.stringify(rows));
  });

  it('encodes and decodes larger tabular data', () => {
    const rows = [
      { name: 'A', qty: 1, rate: 10, amount: 10 },
      { name: 'B', qty: 2, rate: 10, amount: 20 },
      { name: 'C', qty: 3, rate: 10, amount: 30 },
    ];
    const encoded = toToon(rows);
    expect(encoded.length).toBeGreaterThan(0);
    expect(fromToon(encoded)).toEqual(rows);
  });

  it('estimates tokens by 4-char chunks', () => {
    expect(estimateTokens('')).toBe(0);
    expect(estimateTokens('1234')).toBe(1);
    expect(estimateTokens('12345')).toBe(2);
  });
});
