import { describe, it, expect } from 'vitest';
import { resolveListSource } from './validation.handler';

/**
 * TASKS.md #166 — the two Office.js traps a DATA_VALIDATION caller falls into.
 *
 * Both fail SILENTLY: the rule installs, no error is raised, and the dropdown
 * simply matches nothing. That is the same failure shape as #156/#157 — a
 * valid-looking action that delivers nothing — so it is encoded here rather
 * than left to the model to remember.
 */
describe('resolveListSource', () => {
  it('joins literal values into one string, not an array', () => {
    // Passing the array straight through is the natural guess and produces a
    // rule that matches nothing.
    expect(resolveListSource(['Paid', 'Pending', 'Partial'])).toBe('Paid,Pending,Partial');
  });

  it('trims and drops empty literal entries', () => {
    expect(resolveListSource([' Paid ', '', 'Pending'])).toBe('Paid,Pending');
  });

  it('adds the leading = a range reference needs', () => {
    // A planner writing the plain reference is right about intent and wrong
    // about syntax; Office.js needs the `=`.
    expect(resolveListSource('Lists!$B$3:$B$20')).toBe('=Lists!$B$3:$B$20');
  });

  it('leaves an already-prefixed reference alone', () => {
    expect(resolveListSource('=Lists!$B$3:$B$20')).toBe('=Lists!$B$3:$B$20');
  });

  it('does not mistake a comma-separated literal string for a reference', () => {
    expect(resolveListSource('Paid,Pending')).toBe('Paid,Pending');
  });

  it('handles a missing source without throwing', () => {
    expect(resolveListSource(undefined)).toBe('');
    expect(resolveListSource('')).toBe('');
  });

  it('recognises a reference to a sheet whose name has no quotes', () => {
    expect(resolveListSource('Lists!B3:B20')).toBe('=Lists!B3:B20');
  });
});
