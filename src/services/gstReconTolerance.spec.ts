import { describe, expect, it } from 'vitest';
import { toleranceFromMessage } from './gstReconChat';

describe('P2.3 recon tolerance from the prompt', () => {
  it.each([
    ['Reconcile purchase register with GSTR-2B with ₹100 tolerance', 100],
    ['Reconcile with GSTR-2B, tolerance of Rs 50', 50],
    ['reconcile 2b tolerance 1', 1],
    ['Reconcile purchase register with GSTR-2B', null],
  ])('%s → %s', (m, v) => expect(toleranceFromMessage(m)).toBe(v));
});
