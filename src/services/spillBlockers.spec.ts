import { describe, expect, it } from 'vitest';
import {
  buildSpillBlockages,
  describeCells,
  spillBlockersFromGrid,
  spillWidthFromHeader,
} from './spillBlockers';
import { describeOutcome, OutcomeVerification } from './outcomeVerifier';
import { buildRepairRequest } from './repairRequest';

/**
 * TASKS.md #321 — the live case: Main!A19 holds the consolidated-view formula,
 * spilling under a 14-column header at A18:N18, and the build also wrote
 * formulas into G19, I19 and M19. The card said only "1 formula cell(s)
 * returned an error: Main!A19 #SPILL!".
 */
const HEADER = [
  'Month', 'Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out', 'Nights',
  'Rate Per Night', 'Total Amount', 'Source', 'Payment Status', 'Amount Received',
  'Balance Due', 'Bank Account', '', '',
];

describe('spill area inference', () => {
  it('takes its width from the contiguous header labels above the formula', () => {
    expect(spillWidthFromHeader(HEADER)).toBe(14);
    expect(spillWidthFromHeader(['', 'x'])).toBe(0);
  });

  it('names cells holding a formula even when it shows "" — the live G19 case', () => {
    const width = 14;
    const values: unknown[][] = [Array.from({ length: width }, (_, c) => (c === 0 ? '#SPILL!' : ''))];
    const formulas: unknown[][] = [Array.from({ length: width }, () => '')];
    formulas[0][0] = '=LET(rows,VSTACK(...))';
    formulas[0][6] = '=IF(OR(E19="",F19=""),"",F19-E19)'; // G19 → ""
    formulas[0][8] = '=IF(OR(G19="",H19=""),"",G19*H19)'; // I19 → ""
    formulas[0][12] = '=IF(I19="","",I19-N(L19))'; // M19 → ""

    expect(spillBlockersFromGrid(values, formulas, { row: 18, col: 0 })).toEqual([
      'G19',
      'I19',
      'M19',
    ]);
  });

  it('never counts the formula cell itself as blocking', () => {
    expect(spillBlockersFromGrid([['#SPILL!']], [['=SEQUENCE(3)']], { row: 0, col: 0 })).toEqual([]);
  });

  it('describes cells readably and bounds long lists', () => {
    expect(describeCells(['G19'])).toBe('G19');
    expect(describeCells(['G19', 'I19', 'M19'])).toBe('G19, I19 and M19');
    expect(describeCells(['A1', 'B1', 'C1', 'D1', 'E1', 'F1'])).toBe('A1, B1, C1, D1 and 2 more');
  });
});

describe('a blocked spill in the post-apply report', () => {
  const verification: OutcomeVerification = {
    verified: 69,
    unreadable: 0,
    skipped: false,
    mismatches: [
      {
        sheet: 'Main',
        cell: 'A19',
        expected: '',
        actual: '#SPILL!',
        isFormulaError: true,
        formula: '=LET(rows,VSTACK(...))',
        spillBlockers: ['G19', 'I19', 'M19'],
      },
    ],
  };

  it('says what is in the way instead of "returned an error"', () => {
    const message = describeOutcome(verification)!;
    expect(message).toContain('Main!A19');
    expect(message).toContain('G19, I19 and M19');
    expect(message).not.toMatch(/returned an error/);
  });

  it('is not sent to the model as a formula to rewrite — the formula is correct', () => {
    expect(buildRepairRequest(verification)).toBeNull();
  });

  it('becomes a one-click clear naming exactly those cells', () => {
    expect(buildSpillBlockages(verification.mismatches)).toEqual([
      { sheet: 'Main', anchor: 'A19', blockers: ['G19', 'I19', 'M19'] },
    ]);
  });

  it('keeps the plain report when the blockers could not be worked out', () => {
    const unknown: OutcomeVerification = {
      ...verification,
      mismatches: [{ ...verification.mismatches[0], spillBlockers: undefined }],
    };
    expect(describeOutcome(unknown)).toMatch(/returned an error: Main!A19 #SPILL!/);
    expect(buildSpillBlockages(unknown.mismatches)).toEqual([]);
  });
});
