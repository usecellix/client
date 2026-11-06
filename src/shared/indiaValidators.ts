/**
 * Client copy of the GSTIN / PAN checks in Server/src/excel-ai/utils/india-validators.util.ts
 * (B10). The add-in cannot import server code, so the algorithm is duplicated here and pinned by
 * client/src/shared/indiaValidators.spec.ts, which uses the same fixtures as the server spec.
 * Change both together. Classifies only; never "fixes" an identifier.
 */

/** State / UT codes accepted in a GSTIN (TODO(B5): move to the rule-table registry). */
const GST_STATE_CODE_SET = new Set([
  ...Array.from({ length: 38 }, (_, i) => String(i + 1).padStart(2, '0')),
  '97',
  '99',
]);

const PAN_ENTITY_LETTERS = new Set(['P', 'C', 'H', 'F', 'A', 'B', 'T', 'L', 'J', 'G']);
const GSTIN_FORMAT = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;
const CHARSET = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';

export function normalizeIdentifier(raw: unknown): string {
  return String(raw ?? '')
    .replace(/[\s\u00A0]+/g, '')
    .toUpperCase();
}

export function gstinCheckChar(first14: string): string | null {
  if (first14.length !== 14) return null;
  let sum = 0;
  for (let i = 0; i < 14; i += 1) {
    const value = CHARSET.indexOf(first14[i]);
    if (value < 0) return null;
    const product = value * (i % 2 === 0 ? 1 : 2);
    sum += Math.floor(product / 36) + (product % 36);
  }
  return CHARSET[(36 - (sum % 36)) % 36];
}

/** True only when format, state code, PAN segment and check character are all right. */
export function isValidGstin(raw: unknown): boolean {
  const gstin = normalizeIdentifier(raw);
  if (!GSTIN_FORMAT.test(gstin)) return false;
  if (!GST_STATE_CODE_SET.has(gstin.slice(0, 2))) return false;
  if (!PAN_ENTITY_LETTERS.has(gstin[5])) return false; // 4th PAN char = 6th GSTIN char
  return gstinCheckChar(gstin.slice(0, 14)) === gstin[14];
}
