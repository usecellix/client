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
    logApplyStart: vi.fn(),
    logAcceptSuccess: vi.fn(),
    logAcceptFail: vi.fn(),
    logReject: vi.fn(),
    logPreviewStart: vi.fn(),
    logPreviewFail: vi.fn(),
    logAction: vi.fn(),
  },
}));

// The panel is App's whole output; the props App hands it are kept so a test can call them.
const panel = vi.hoisted(() => ({ props: null as null | Record<string, unknown> }));
vi.mock('@/components/ConversationPanel/ConversationPanel', () => ({
  default: (props: Record<string, unknown>) => {
    panel.props = props;
    return <div data-testid="app-rendered" />;
  },
}));

import App from '@/taskpane/App';
import { useConversation } from '@/hooks/useConversation';
import { previewManager } from '@/services/previewManager';
import { markChangeSetApplied } from '@/services/auditService';
import { ActionEngine } from '@/utils/actionEngine';

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
      importBankStatement: vi.fn(),
      answerQuestion: vi.fn(),
      answerClarification: vi.fn(),
      dismissClarification: vi.fn(),
      acceptActions: vi.fn(),
    acceptAllActions: vi.fn(),
      rejectActions: vi.fn(),
      resolveGstReconCollisionChoice: vi.fn(),
      rejectGstReconCollisionChoice: vi.fn(),
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
      resumableRun: null,
      resumeRun: vi.fn(),
      dismissResumableRun: vi.fn(),
      retryMissingSheets: vi.fn(),
      markChangeSetReverted: vi.fn(),
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

/**
 * TASKS.md #404 — a revert is recorded on the message once the workbook is back,
 * so the card and the menu can say Reverted after the menu is closed. Recording
 * it before the inverse has applied would hide a Revert that did not happen.
 */
describe('App — a revert is recorded only after it has been applied (TASKS.md #404)', () => {
  afterEach(() => {
    cleanup();
    vi.restoreAllMocks();
    vi.clearAllMocks();
  });

  type RevertHandler = (changeSetId: string, inverseActions: unknown[]) => Promise<void>;
  function renderAndGetRevert() {
    renderAppAndCaptureOptions();
    const onRevert = panel.props?.onRevertHistoryEntry as RevertHandler;
    const results = useConversationMock.mock.results;
    const markReverted = results[results.length - 1]?.value.markChangeSetReverted as ReturnType<typeof vi.fn>;
    return { onRevert, markReverted };
  }
  const inverse = [{ type: 'FORMAT_RANGE', sheetName: 'S', row: 0, col: 0, format: { clearFill: true } }];

  it('applies the inverse actions, then records the change as reverted', async () => {
    const order: string[] = [];
    vi.spyOn(ActionEngine, 'applyActions').mockImplementation(async () => {
      order.push('apply');
    });
    const { onRevert, markReverted } = renderAndGetRevert();
    markReverted.mockImplementation(() => order.push('record'));

    await act(async () => {
      await onRevert('cs_1', inverse);
    });

    expect(order).toEqual(['apply', 'record']);
    expect(markReverted).toHaveBeenCalledWith('cs_1');
  });

  it('does not record it when applying the inverse fails', async () => {
    vi.spyOn(ActionEngine, 'applyActions').mockRejectedValue(new Error('Spreadsheet update failed'));
    const { onRevert, markReverted } = renderAndGetRevert();

    await act(async () => {
      await expect(onRevert('cs_1', inverse)).rejects.toThrow(/Spreadsheet update failed/);
    });

    expect(markReverted).not.toHaveBeenCalled();
  });

  it('records a change that was already reverted elsewhere, with nothing to apply', async () => {
    const apply = vi.spyOn(ActionEngine, 'applyActions').mockResolvedValue(undefined);
    const { onRevert, markReverted } = renderAndGetRevert();

    await act(async () => {
      await onRevert('cs_1', []);
    });

    expect(apply).not.toHaveBeenCalled();
    expect(markReverted).toHaveBeenCalledWith('cs_1');
  });
});
