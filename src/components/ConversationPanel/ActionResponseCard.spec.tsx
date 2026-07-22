import { describe, expect, it, vi } from 'vitest';
import { renderToStaticMarkup } from 'react-dom/server';
import React from 'react';
import { ActionResponseCard } from '@/components/ConversationPanel/ActionResponseCard';
import { ActionBlock } from '@/types/conversationTurn';
import { INTERNAL_COPY_MARKERS, resolveActionBlockCopy } from '@/utils/userFacingResponse';

const FORBIDDEN = [
  'Tier',
  'single-action',
  'no verification',
  'CONDITIONAL_FORMAT',
  'Direct Change',
  'Planner',
  'Verifier',
  'openai/',
];

function makeBlock(overrides: Partial<ActionBlock> = {}): ActionBlock {
  return {
    id: 'actions_1',
    type: 'actions',
    actions: [
      {
        type: 'FORMAT_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A2:L51',
        format: { fillColor: '#FFEB9C' },
      },
    ],
    explanation: 'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
    proposalStatus: 'pending',
    changeSetId: 'cs_1',
    changes: Array.from({ length: 9 }, (_, i) => ({
      cell: `A${53 + Math.floor(i / 3)}${String.fromCharCode(65 + (i % 3))}`,
      sheet: 'Purchase Register',
      before: null,
      after: 1,
      isHardcoded: true,
    })),
    userFacingSummary: {
      contextLine: 'Working with: Purchase Register',
      headline: "I'll highlight Pending payment rows in yellow.",
      supportingDetail: '9 cells, Purchase Register!A53:C55',
    },
    internalDetails: {
      tier: 1,
      model: 'openai/gpt-5-mini',
      processingLabel: 'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
      rawActionSummary: 'FORMAT_MATCHING_ROWS Purchase Register A2:L51',
      legacyExplanation: 'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
    },
    ...overrides,
  };
}

describe('ActionResponseCard', () => {
  it('keeps forbidden internal strings out of the default view', () => {
    const html = renderToStaticMarkup(
      React.createElement(ActionResponseCard, {
        block: makeBlock(),
        previewEnabled: true,
        showActionButtons: true,
        onAccept: vi.fn(),
        onReject: vi.fn(),
        defaultDetailsExpanded: false,
      }),
    );

    const defaultPart = html.split('data-testid="action-details"')[0] ?? html;
    for (const token of FORBIDDEN) {
      expect(defaultPart.includes(token)).toBe(false);
    }
    expect(html).toContain('highlight Pending payment rows in yellow');
    expect(html).toContain('9 cells, Purchase Register!A53:C55');
    expect(html).toContain('Accept');
    expect(html).toContain('Reject');
    expect(html).toContain('Show details');
  });

  it('preserves internals inside the expandable details section', () => {
    const html = renderToStaticMarkup(
      React.createElement(ActionResponseCard, {
        block: makeBlock(),
        previewEnabled: true,
        showActionButtons: true,
        onAccept: vi.fn(),
        onReject: vi.fn(),
        defaultDetailsExpanded: true,
      }),
    );

    expect(html).toContain('data-testid="action-details-body"');
    expect(html).toContain('openai/gpt-5-mini');
    expect(html).toContain('Tier 1 single-action');
    expect(html).toContain('CONDITIONAL_FORMAT');
    expect(html).toContain('FORMAT_MATCHING_ROWS');
  });

  it('surfaces assumptions in the default headline', () => {
    const html = renderToStaticMarkup(
      React.createElement(ActionResponseCard, {
        block: makeBlock({
          userFacingSummary: {
            headline:
              "I noticed 'paid should be first then paid' looks like a typo — I've sorted with Paid first, then Pending.",
          },
          internalDetails: {
            tier: 1,
            assumption: 'Paid first, then Pending',
            processingLabel: 'Tier 1 single-action (SORT_OR_FILTER)',
          },
        }),
        previewEnabled: true,
        onAccept: vi.fn(),
        onReject: vi.fn(),
      }),
    );

    expect(html).toContain('typo');
    expect(html).toContain('Paid first');
    const defaultPart = html.split('data-testid="action-details"')[0] ?? html;
    expect(defaultPart.includes('Tier 1')).toBe(false);
  });
});

describe('resolveActionBlockCopy', () => {
  it('falls back without leaking tier jargon from legacy explanation', () => {
    const copy = resolveActionBlockCopy({
      explanation: 'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
      actions: [{ type: 'SORT_RANGE' }],
      changes: [
        { cell: 'A1', sheet: 'S', before: null, after: 1, isHardcoded: true },
        { cell: 'A2', sheet: 'S', before: null, after: 2, isHardcoded: true },
      ],
    });
    expect(INTERNAL_COPY_MARKERS.test(copy.headline)).toBe(false);
    expect(copy.supportingDetail).toBe('2 cells');
  });
});
