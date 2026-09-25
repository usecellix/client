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

  // TASKS.md #313 — the live message a stuck user saw, verbatim.
  it('rewrites the target-range guard message for the user, dropping model-facing instructions', () => {
    const live =
      'Write blocked: target range A19 already contains data. This action would overwrite existing values. ' +
      'Existing values include: #SPILL!. If you meant to add a new column, use INSERT_COLUMN with position ' +
      '"afterLastColumn" instead of writing into an occupied column. To replace existing content on purpose, ' +
      'the request must explicitly confirm overwrite.';
    const clean = toUserFacingApplyError(live);
    expect(clean).toContain('A19');
    expect(clean).toContain('#SPILL!');
    expect(clean).toMatch(/nothing was changed/i);
    expect(clean).toMatch(/reject/i);
    expect(clean).not.toMatch(/INSERT_COLUMN|afterLastColumn|confirm overwrite/);
  });

  it('maps Excel ItemNotFound-style errors to Accept guidance', () => {
    const clean = toUserFacingApplyError("The requested resource doesn't exist.");
    expect(clean.toLowerCase()).toContain('accept');
    expect(clean).not.toBe("The requested resource doesn't exist.");
  });

  /**
   * TASKS.md #263 — the engine throws `${action.type}: ${message}` at APPLY
   * time, so a prefixed ItemNotFound means a sheet this change set was meant
   * to create is missing. "Click Accept" is useless there (they just did), and
   * the old ordering reported it as a formatting problem instead.
   */
  it('distinguishes an apply-time missing sheet from the preview-time case', () => {
    const applyTime = toUserFacingApplyError("BATCH_SET: The requested resource doesn't exist.");
    expect(applyTime).toMatch(/never created/i);
    expect(applyTime).not.toMatch(/formatting/i);

    // Unprefixed is the preview path — Accept really is the right advice.
    const previewTime = toUserFacingApplyError("The requested resource doesn't exist.");
    expect(previewTime.toLowerCase()).toContain('accept');
    expect(previewTime).not.toMatch(/never created/i);
  });
});
