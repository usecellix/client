import { describe, expect, it } from 'vitest';
import { shouldAcceptIncomingClarification } from './clarification.util';
import { ClarificationPayload } from '@/types/cellix.types';

describe('shouldAcceptIncomingClarification', () => {
  const payload: ClarificationPayload = {
    question: 'Which column?',
    suggestions: ['A', 'B'],
    ambiguityScore: 70,
  };

  it('accepts when no clarification is pending', () => {
    expect(shouldAcceptIncomingClarification(null)).toBe(true);
  });

  it('rejects when clarification is already pending', () => {
    expect(shouldAcceptIncomingClarification(payload)).toBe(false);
  });
});
