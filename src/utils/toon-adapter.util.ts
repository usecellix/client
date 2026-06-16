import { decode, encode } from '@toon-format-cjs/toon';

/**
 * Converts a uniform array of objects to TOON format string.
 * Only use when: array has 3+ items AND each item has 4+ fields.
 * Below those thresholds, header overhead makes TOON larger than JSON.
 */
export function toToon(rows: Record<string, unknown>[]): string {
  if (!rows || rows.length === 0) return '[]';

  const firstItem = rows[0];
  const fieldCount = Object.keys(firstItem).length;

  if (rows.length < 3 || fieldCount < 4) {
    return JSON.stringify(rows);
  }

  try {
    return encode(rows);
  } catch {
    return JSON.stringify(rows);
  }
}

/**
 * Converts TOON string back to object array.
 */
export function fromToon(toonStr: string): Record<string, unknown>[] {
  const trimmed = toonStr?.trim() ?? '';
  if (!trimmed) return [];

  try {
    const parsed = decode(trimmed);
    return Array.isArray(parsed) ? (parsed as Record<string, unknown>[]) : [];
  } catch {
    return JSON.parse(trimmed) as Record<string, unknown>[];
  }
}

/**
 * Returns a rough token count estimate (1 token ~= 4 chars).
 */
export function estimateTokens(str: string): number {
  return Math.ceil(str.length / 4);
}
