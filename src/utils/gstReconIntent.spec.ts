import { describe, expect, it } from 'vitest';
import { detectGstReconIntent, isGstReconPrompt } from './gstReconIntent';

describe('gstReconIntent', () => {
  it('detects purchase vs GSTR-2B prompts', () => {
    expect(detectGstReconIntent('reconcile my GST purchases with GSTR-2B')?.type).toBe(
      'PR_VS_GSTR2B',
    );
    expect(isGstReconPrompt('Match purchase register to 2B')).toBe(true);
  });

  it('detects 2A and IMS', () => {
    expect(detectGstReconIntent('reconcile PR with GSTR-2A')?.type).toBe('PR_VS_GSTR2A');
    expect(detectGstReconIntent('IMS vs purchase register reconciliation')?.type).toBe(
      'IMS_VS_PR',
    );
  });

  it('ignores non-GST chat', () => {
    expect(detectGstReconIntent('create a new sheet called Summary')).toBeNull();
    expect(detectGstReconIntent('what is the total in column B?')).toBeNull();
  });
});
