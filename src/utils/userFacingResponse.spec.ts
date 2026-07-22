import { describe, expect, it } from 'vitest';
import {
  INTERNAL_COPY_MARKERS,
  resolveActionBlockCopy,
} from '@/utils/userFacingResponse';

describe('userFacingResponse', () => {
  it('prefers backend userFacingSummary', () => {
    const copy = resolveActionBlockCopy({
      userFacingSummary: {
        headline: 'Added a Payment Status summary.',
        supportingDetail: '9 cells, A53:C55',
      },
      explanation: 'Tier 1 single-action (CONDITIONAL_FORMAT)',
      actions: [{ type: 'AGGREGATE_TABLE' }],
    });
    expect(copy.headline).toBe('Added a Payment Status summary.');
    expect(copy.supportingDetail).toBe('9 cells, A53:C55');
  });

  it('strips legacy tier jargon and uses cell count', () => {
    const copy = resolveActionBlockCopy({
      explanation: 'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
      actions: [{ type: 'FORMAT_MATCHING_ROWS' }],
      changes: Array.from({ length: 9 }, (_, i) => ({
        cell: `A${i + 1}`,
        sheet: 'S',
        before: null,
        after: 1,
        isHardcoded: true,
      })),
    });
    expect(INTERNAL_COPY_MARKERS.test(copy.headline)).toBe(false);
    expect(copy.supportingDetail).toBe('9 cells');
  });
});
