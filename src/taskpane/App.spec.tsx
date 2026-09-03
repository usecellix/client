// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, cleanup, render } from '@testing-library/react';
import type { UseConversationOptions } from '@/hooks/useConversation';

/**
 * App.tsx-level regression coverage for TASKS.md #44/#45 — the false-success
 * pattern where a failed `markChangeSetApplied` audit call must roll the
 * changeSetId back out of `appliedChangeSetIdsRef`, not leave the UI showing
 * "Applied" for a change the audit trail never recorded.
 *
 * The PreviewSummaryBar UI `handlePreviewAccept` originally guarded (App's
 * second, now-removed accept path) was deleted as dead code once that bar was
 * removed in favor of the single per-block Accept on `ActionResponseCard`.
 * This spec now exercises the one remaining accept path directly:
 * `applyActionsWithAudit`, passed to `useConversation` as `onActions` and
 * invoked by the real `acceptActions()` flow.
 *
 * `App.tsx` has no other test coverage — this file builds the minimum harness
 * needed for this regression rather than a general-purpose App render
 * fixture: `useConversation`, `previewManager`, `markChangeSetApplied`, and
 * `frontendTelemetry` are all mocked so the test exercises only
 * `applyActionsWithAudit`'s real, unmocked logic. `ConversationPanel` (App's
 * entire render output) is mocked down to a single button wired to
 * `props.onAcceptActions`, since rendering the real component tree is not
 * what this regression needs and would pull in unrelated complexity.
 *
 * `applyActionsWithAudit` (App's `onActions` callback, passed to
 * `useConversation`) is captured from the mock's call args and invoked
 * directly — this runs App's real, unmocked closure exactly as `acceptActions`
 * would call it, without needing a full SSE round trip.
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
  default: () => <div data-testid="app-rendered" />,
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
    acceptAllActions: vi.fn(),
      rejectActions: vi.fn(),
      endConversation: vi.fn(),
      newChat: vi.fn(),
      clearConversation: vi.fn(),
      selectSession: vi.fn(),
      closeSession: vi.fn(),
      renameSession: vi.fn(),
      deleteSession: vi.fn().mockResolvedValue(undefined),
      deleteHistoryConversation: vi.fn().mockResolvedValue(undefined),
      openConversationFromHistory: vi.fn().mockResolvedValue(true),
      isLoadingHistoryConversation: false,
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

describe('App — applyActionsWithAudit accept-path rollback (TASKS.md #44/#45)', () => {
  beforeEach(() => {
    previewManagerMock.active = false;
    previewManagerMock.accept.mockReset();
    markChangeSetAppliedMock.mockReset();
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  const sampleActions = [{ type: 'SET_CELL', address: 'A1', value: 5 }] as never;

  it('rolls back appliedChangeSetIds when markChangeSetApplied fails on accept', async () => {
    const options = renderAppAndCaptureOptions();
    previewManagerMock.active = true;

    previewManagerMock.accept.mockResolvedValue({
      createdConditionalFormatIds: [],
      createdChartIds: [],
    } as never);
    markChangeSetAppliedMock.mockRejectedValueOnce(new Error('audit sync failed'));

    // Real path: `acceptActions()` (mocked here) calls App's `onActions`
    // (`applyActionsWithAudit`) exactly like this — captured and invoked
    // directly so the test doesn't need a full turns/SSE round trip.
    await act(async () => {
      await expect(
        options.onActions?.(sampleActions, 'Set A1 to 5.', {
          changeSetId: CHANGE_SET_ID,
          changes: [],
        } as never),
      ).rejects.toThrow();
    });

    // #44's fix: a failed markChangeSetApplied must roll the id back out of
    // appliedChangeSetIdsRef, not leave it permanently marked applied.
    expect(options.isChangeSetApplied?.(CHANGE_SET_ID)).toBe(false);
  });

  it('marks the changeSetId applied when markChangeSetApplied succeeds (sanity check)', async () => {
    const options = renderAppAndCaptureOptions();
    previewManagerMock.active = true;

    previewManagerMock.accept.mockResolvedValue({
      createdConditionalFormatIds: [],
      createdChartIds: [],
    } as never);
    markChangeSetAppliedMock.mockResolvedValueOnce({} as never);

    await act(async () => {
      await options.onActions?.(sampleActions, 'Set A1 to 5.', {
        changeSetId: CHANGE_SET_ID,
        changes: [],
      } as never);
    });

    expect(options.isChangeSetApplied?.(CHANGE_SET_ID)).toBe(true);
  });
});
