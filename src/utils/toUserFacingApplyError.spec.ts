import { describe, expect, it } from 'vitest';
import { toUserFacingApplyError } from '@/utils/toUserFacingApplyError';

describe('toUserFacingApplyError (Spec 24)', () => {
  it('sanitizes FORMAT_MATCHING_ROWS validation leak', () => {
    const clean = toUserFacingApplyError(
      'FORMAT_MATCHING_ROWS: findMatchingRowOffsets requires hasHeaders: true',
    );
    expect(clean).not.toMatch(/FORMAT_MATCHING_ROWS|findMatchingRowOffsets|hasHeaders/i);
    expect(clean.toLowerCase()).toContain("couldn't apply");
  });

  it('sanitizes Spreadsheet update failed wrappers', () => {
    const clean = toUserFacingApplyError(
      'Spreadsheet update failed: FORMAT_MATCHING_ROWS: findMatchingRowOffsets requires hasHeaders: true',
    );
    expect(clean).not.toMatch(/Spreadsheet update failed|FORMAT_MATCHING/i);
  });

  it('preserves overwrite guard copy', () => {
    const msg = 'Write blocked: target range K2 is occupied';
    expect(toUserFacingApplyError(msg)).toBe(msg);
  });

  it('maps Excel ItemNotFound-style errors to Accept guidance', () => {
    const clean = toUserFacingApplyError("The requested resource doesn't exist.");
    expect(clean.toLowerCase()).toContain('accept');
    expect(clean).not.toBe("The requested resource doesn't exist.");
  });
});
