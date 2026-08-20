// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { UseConversationOptions } from '@/hooks/useConversation';

/**
 * App.tsx-level regression coverage for TASKS.md #44/#45 — `handlePreviewAccept`'s
 * "no pending action block, previewManager.active only" path (App.tsx:278-301).
 *
 * `App.tsx` has no other test coverage — this file builds the minimum harness
 * needed for this one regression rather than a general-purpose App render
 * fixture: `useConversation`, `previewManager`, `markChangeSetApplied`, and
 * `frontendTelemetry` are all mocked so the test exercises only
 * `handlePreviewAccept`'s real, unmocked logic. `ConversationPanel` (App's
 * entire render output) is mocked down to a single button wired to
 * `props.pendingPreview.onAccept`, since rendering the real component tree is
 * not what this regression needs and would pull in unrelated complexity.
 *
 * `previewActions` (App's `onPreviewActions` callback, passed to
 * `useConversation`) is captured from the mock's call args and invoked
 * directly — this runs App's real, unmocked `previewActions` closure exactly
 * as the SSE 'actions' handler would call it, without needing a full SSE
 * round trip.
 */

vi.mock('@/hooks/useConversation', () => ({
  useConversation: vi.fn(),
}));

vi.mock('@/services/previewManager', () => ({
  previewManager: {
    active: false,
    render: vi.fn().mockResolvedValue([]),
    highlightChanges: vi.fn().mockResolvedValue(undefined),
    accept: vi.fn(),
    reject: vi.fn().mockResolvedValue(undefined),
  },
}));

vi.mock('@/services/auditService', () => ({
  markChangeSetApplied: vi.fn(),
}));

vi.mock('@/services/frontendTelemetry', () => ({
  frontendTelemetry: {
    installConsoleCapture: vi.fn(),
    setContext: vi.fn(),
    logAcceptClick: vi.fn(),
    logAcceptSuccess: vi.fn(),
    logAcceptFail: vi.fn(),
    logReject: vi.fn(),
    logPreviewStart: vi.fn(),
    logPreviewFail: vi.fn(),
    logAction: vi.fn(),
  },
}));

vi.mock('@/components/ConversationPanel/ConversationPanel', () => ({
  default: (props: { pendingPreview?: { onAccept: () => void } | null }) =>
    props.pendingPreview ? (
      <button type="button" data-testid="accept-preview" onClick={() => props.pendingPreview!.onAccept()}>
        Accept
      </button>
    ) : (
      <div data-testid="no-preview" />
    ),
}));

import App from '@/taskpane/App';
import { useConversation } from '@/hooks/useConversation';
import { previewManager } from '@/services/previewManager';
import { markChangeSetApplied } from '@/services/auditService';

const useConversationMock = vi.mocked(useConversation);
const previewManagerMock = vi.mocked(previewManager, true);
const markChangeSetAppliedMock = vi.mocked(markChangeSetApplied);

const CHANGE_SET_ID = 'cs_race_1';

function renderAppAndCaptureOptions(): UseConversationOptions {
  let captured: UseConversationOptions | undefined;
  useConversationMock.mockImplementation((options) => {
    captured = options;
    return {
      sessions: [],
      activeSessionId: null,
      turns: [], // No pending ActionBlock in any turn — forces the previewManager-active-only branch.
      activeTurnId: null,
      isWaitingForResponse: false,
      isWaitingClarification: false,
      activeClarification: null,
      conversationId: null,
      sendMessage: vi.fn(),
      answerQuestion: vi.fn(),
      answerClarification: vi.fn(),
      dismissClarification: vi.fn(),
      acceptActions: vi.fn(),
      rejectActions: vi.fn(),
      endConversation: vi.fn(),
      newChat: vi.fn(),
      clearConversation: vi.fn(),
      selectSession: vi.fn(),
      closeSession: vi.fn(),
      selectTurn: vi.fn(),
      closeTurn: vi.fn(),
      toggleThinking: vi.fn(),
      markAnswerComplete: vi.fn(),
    };
  });

  render(<App />);
  if (!captured) throw new Error('useConversation was not called during render');
  return captured;
}

describe('App — handlePreviewAccept (TASKS.md #44/#45)', () => {
  beforeEach(() => {
    previewManagerMock.active = false;
    previewManagerMock.accept.mockReset();
    markChangeSetAppliedMock.mockReset();
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it('rolls back appliedChangeSetIds when markChangeSetApplied fails on the previewManager-active-only path', async () => {
    const options = renderAppAndCaptureOptions();

    // Simulate the real SSE handler calling onPreviewActions with no matching
    // turn block yet in `turns` (mocked empty) — sets hasPendingPreview +
    // pendingChangeSetId via App's real, unmocked previewActions closure.
    await act(async () => {
      await options.onPreviewActions?.('Set A1 to 5.' as never, 'Set A1 to 5.', {
        changeSetId: CHANGE_SET_ID,
        changes: [],
      } as never);
    });
    previewManagerMock.active = true;

    previewManagerMock.accept.mockResolvedValue({
      createdConditionalFormatIds: [],
      createdChartIds: [],
    } as never);
    markChangeSetAppliedMock.mockRejectedValueOnce(new Error('audit sync failed'));

    const acceptButton = await screen.findByTestId('accept-preview');
    await act(async () => {
      fireEvent.click(acceptButton);
    });

    // #44's fix: a failed markChangeSetApplied must roll the id back out of
    // appliedChangeSetIdsRef, not leave it permanently marked applied.
    expect(options.isChangeSetApplied?.(CHANGE_SET_ID)).toBe(false);
  });

  it('marks the changeSetId applied when markChangeSetApplied succeeds (sanity check)', async () => {
    const options = renderAppAndCaptureOptions();

    await act(async () => {
      await options.onPreviewActions?.('Set A1 to 5.' as never, 'Set A1 to 5.', {
        changeSetId: CHANGE_SET_ID,
        changes: [],
      } as never);
    });
    previewManagerMock.active = true;

    previewManagerMock.accept.mockResolvedValue({
      createdConditionalFormatIds: [],
      createdChartIds: [],
    } as never);
    markChangeSetAppliedMock.mockResolvedValueOnce({} as never);

    const acceptButton = await screen.findByTestId('accept-preview');
    await act(async () => {
      fireEvent.click(acceptButton);
    });

    expect(options.isChangeSetApplied?.(CHANGE_SET_ID)).toBe(true);
  });
});
