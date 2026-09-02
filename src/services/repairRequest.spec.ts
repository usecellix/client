import { describe, it, expect } from 'vitest';
import { buildRepairRequest } from './repairRequest';
import { OutcomeVerification, OutcomeMismatch } from './outcomeVerifier';

/**
 * TASKS.md #168 — the read-back can now produce a fix, not just a complaint.
 *
 * The scenario is taken from Shortcut's own planning log part 6: its
 * consolidation formula returned `#REF!` because a `LET` variable name collided
 * with a real range reference. It saw the error and repaired itself; #150 gave
 * us the seeing and stopped there.
 */
function verification(mismatches: OutcomeMismatch[]): OutcomeVerification {
  return { verified: 0, mismatches, unreadable: 0, skipped: false };
}

const refError: OutcomeMismatch = {
  sheet: 'Main',
  cell: 'A19',
  expected: '',
  actual: '#REF!',
  isFormulaError: true,
  formula: '=LET(all,VSTACK(January!A2:J500),all)',
};

describe('buildRepairRequest', () => {
  it('returns null when nothing went wrong', () => {
    expect(buildRepairRequest(verification([]))).toBeNull();
  });

  it('returns null when verification could not run', () => {
    // Unknown is not the same as clean — never fabricate a repair from it.
    expect(
      buildRepairRequest({ ...verification([refError]), skipped: true }),
    ).toBeNull();
  });

  it('builds a repair naming the cell, the error and the formula', () => {
    const repair = buildRepairRequest(verification([refError]));
    expect(repair).not.toBeNull();
    expect(repair!.prompt).toContain('Main!A19');
    expect(repair!.prompt).toContain('#REF!');
    // The formula is the CAUSE; `expected` holds only the recorded value.
    expect(repair!.prompt).toContain('=LET(all,VSTACK(January!A2:J500),all)');
    expect(repair!.cellCount).toBe(1);
    expect(repair!.errors).toEqual(['#REF!']);
  });

  it('collapses one broken formula repeated down many rows into a single item', () => {
    // A build that writes the same bad formula on 12 month rows has ONE bug.
    const many = Array.from({ length: 12 }, (_, i) => ({
      ...refError,
      cell: `B${i + 5}`,
      formula: '=SUM(NoSuchSheet!G:G)',
    }));
    const repair = buildRepairRequest(verification(many))!;
    expect(repair.cellCount).toBe(12);
    // One bullet, not twelve.
    expect(repair.prompt.split('\n').filter((l) => l.startsWith('- '))).toHaveLength(1);
    expect(repair.prompt).toContain('4 more cell(s)');
  });

  it('separates genuinely different failures', () => {
    const repair = buildRepairRequest(
      verification([
        refError,
        { ...refError, cell: 'B7', actual: '#NAME?', formula: '=BADFN(1)' },
      ]),
    )!;
    expect(repair.prompt.split('\n').filter((l) => l.startsWith('- '))).toHaveLength(2);
    expect(repair.errors.sort()).toEqual(['#NAME?', '#REF!']);
  });

  it('ignores plain value mismatches — Excel smart-entry is not a bug to fight', () => {
    const smartEntry: OutcomeMismatch = {
      sheet: 'January',
      cell: 'D2',
      expected: '12-09-26',
      actual: '120926',
      isFormulaError: false,
    };
    expect(buildRepairRequest(verification([smartEntry]))).toBeNull();
  });

  it('ignores a missing sheet — that needs a rebuild, not a formula fix', () => {
    const missing: OutcomeMismatch = {
      sheet: 'Ghost',
      cell: 'A1',
      expected: 'x',
      actual: '(sheet missing)',
      isFormulaError: true,
    };
    expect(buildRepairRequest(verification([missing]))).toBeNull();
  });

  it('still builds a repair when the formula text was not captured', () => {
    const noFormula = { ...refError, formula: undefined };
    const repair = buildRepairRequest(verification([noFormula]))!;
    expect(repair.prompt).toContain('Main!A19');
    expect(repair.prompt).not.toContain('The formula written there was');
  });

  it('tells the model to fix only the failing cells', () => {
    const repair = buildRepairRequest(verification([refError]))!;
    expect(repair.prompt).toContain('leave everything else exactly as it is');
  });
});
