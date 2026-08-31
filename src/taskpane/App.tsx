import React, { useCallback, useEffect, useRef, useState } from 'react';
import ConversationPanel from '@/components/ConversationPanel/ConversationPanel';
import { CompareResult } from '@/components/SheetCompareView/SheetCompareView';
import { useConversation, PreviewActionsMeta } from '@/hooks/useConversation';
import { ActionEngine } from '@/utils/actionEngine';
import type { CreatedConditionalFormatId, CreatedChartId } from '@/engine/actionEngine';
import { CellChange } from '@/types/changeSet';
import { previewManager } from '@/services/previewManager';
import { markChangeSetApplied } from '@/services/auditService';
import {
  describeOutcome,
  verifyAppliedOutcomeSafe,
} from '@/services/outcomeVerifier';
import { frontendTelemetry } from '@/services/frontendTelemetry';
import {
  getContextForSend,
  markPendingWorkbookContextStale,
} from '@/utils/pendingWorkbookContext';
import { SheetAction } from '@/types/sheet-actions';
import { RestoreResult } from '@/types/checkpoint';
import { AssistantMode, DEFAULT_ASSISTANT_MODE, isAssistantMode } from '@/types/mode';
import { resolveWorkbookKey, loadChatSessions, saveChatSessions } from '@/utils/chatSessionStorage';
import { resolveWorkbookId } from '@/utils/workbookIdentity';
import '@/styles/conversation-panel.css';
import './taskpane.css';

/* global Excel */

const App: React.FC = () => {
  const previewEnabled = true;
  const [mode, setMode] = useState<AssistantMode>(DEFAULT_ASSISTANT_MODE);
  const [isApplying, setIsApplying] = useState(false);
  const [compareResult, setCompareResult] = useState<CompareResult | null>(null);
  const [isComparing] = useState(false);
  const [isReadingWorkbook, setIsReadingWorkbook] = useState(false);
  const applyInProgressRef = useRef(false);
  const appliedChangeSetIdsRef = useRef<Set<string>>(new Set());
  const [workbookKey, setWorkbookKey] = useState('workbook');
  // Durable per-workbook identity (TASKS.md #22-23), minted/persisted via
  // Office.js document.settings — threaded into useConversation below so it
  // rides along on every conversation-creating request.
  const [workbookId, setWorkbookId] = useState<string | undefined>(undefined);

  const isChangeSetApplied = useCallback((changeSetId?: string) => {
    return Boolean(changeSetId && appliedChangeSetIdsRef.current.has(changeSetId));
  }, []);

  const handleModeChange = useCallback(
    (next: AssistantMode) => {
      setMode(next);
      if (!workbookKey) return;
      const existing = loadChatSessions(workbookKey);
      saveChatSessions(workbookKey, {
        activeSessionId: existing?.activeSessionId ?? null,
        sessions: existing?.sessions ?? [],
        assistantMode: next,
      });
    },
    [workbookKey],
  );

  const applyActionsWithAudit = useCallback(
    async (actions: SheetAction[], explanation: string, meta?: PreviewActionsMeta) => {
      if (!actions.length) return;

      frontendTelemetry.logAcceptClick(actions, {
        changeSetId: meta?.changeSetId,
        source: 'applyActionsWithAudit',
      });

      let createdConditionalFormatIds: CreatedConditionalFormatId[] | undefined;
      let createdChartIds: CreatedChartId[] | undefined;
      let sortedRangeChanges: CellChange[] | undefined;
      try {
        if (previewManager.active) {
          const result = await previewManager.accept();
          createdConditionalFormatIds = result?.createdConditionalFormatIds;
          createdChartIds = result?.createdChartIds;
          sortedRangeChanges = result?.sortedRangeChanges;
        } else if (meta?.changeSetId && appliedChangeSetIdsRef.current.has(meta.changeSetId)) {
          // Already applied earlier — do not re-run INSERT_COLUMN / writes.
        } else {
          // Preview never started or was cleared — apply directly. Report (not
          // throw) errors, matching applyActions' own throw condition below —
          // TASKS.md #40/#15 need the created-id lists from this path too.
          const result = await ActionEngine.applyActionsWithReport(actions);
          if (result.errors.length > 0 && result.applied === 0) {
            throw new Error(result.errors.join('; '));
          }
          createdConditionalFormatIds = result.createdConditionalFormatIds;
          createdChartIds = result.createdChartIds;
          sortedRangeChanges = result.sortedRangeChanges;
        }

        if (meta?.changeSetId) {
          appliedChangeSetIdsRef.current.add(meta.changeSetId);
          try {
            await markChangeSetApplied(
              meta.changeSetId,
              createdConditionalFormatIds,
              createdChartIds,
              sortedRangeChanges,
            );
          } catch (error) {
            // Spec 22 Bug 3: do not swallow apply failures — UI must not show Applied.
            appliedChangeSetIdsRef.current.delete(meta.changeSetId);
            const message =
              error instanceof Error
                ? error.message
                : 'This change could not be applied — audit sync failed. Try again?';
            frontendTelemetry.logAcceptFail(error, actions, { changeSetId: meta.changeSetId });
            throw new Error(message);
          }
        }

        // Spec 09 item 1: change-set apply invalidates the pending workbook prebuild.
        markPendingWorkbookContextStale();

        // TASKS.md #150 — outcome verification. Everything above this line
        // checked INTENT (were the actions well-formed, did the engine report
        // success). This is the only place that checks the WORKBOOK: read back
        // the cells the ChangeSet claims to have written and compare against
        // its own recorded `after` values. `CODEBASE_ANALYSIS.md` §3.15.
        //
        // Never throws and never blocks: the write already happened, so a
        // read-back problem must not turn a real success into a failure. It
        // reports, and reporting honestly is the entire point (§3.7).
        const verification = await verifyAppliedOutcomeSafe(meta?.changes ?? []);
        const outcomeMessage = describeOutcome(verification);
        if (outcomeMessage) {
          console.warn('[Cellix] Post-apply verification found problems:', verification);
          frontendTelemetry.logAction(
            'verify',
            'verify.mismatch',
            outcomeMessage,
            {
              changeSetId: meta?.changeSetId,
              verified: verification.verified,
              mismatchCount: verification.mismatches.length,
              unreadable: verification.unreadable,
              sample: verification.mismatches.slice(0, 5),
            },
          );
        } else if (!verification.skipped && verification.verified > 0) {
          frontendTelemetry.logAction(
            'verify',
            'verify.ok',
            `Verified ${verification.verified} written cell(s) against the workbook`,
            { changeSetId: meta?.changeSetId, verified: verification.verified },
          );
        }
        meta?.onOutcomeVerified?.(verification, outcomeMessage);

        frontendTelemetry.logAcceptSuccess(actions, {
          changeSetId: meta?.changeSetId,
          explanation,
        });
        console.info('[Cellix] Applied actions:', explanation, actions);
      } catch (error) {
        frontendTelemetry.logAcceptFail(error, actions, { changeSetId: meta?.changeSetId });
        throw error;
      }
    },
    [],
  );

  const previewActions = useCallback(
    async (actions: SheetAction[], explanation: string, meta?: PreviewActionsMeta) => {
      if (!actions.length) return;

      const changes = meta?.changes ?? [];

      frontendTelemetry.logPreviewStart(actions, explanation, { changeSetId: meta?.changeSetId });

      try {
        if (changes.length > 0) {
          await previewManager.highlightChanges(changes, actions);
        } else {
          await previewManager.render({ actions, summary: explanation });
        }
        frontendTelemetry.logAction(
          'preview',
          'preview.ready',
          // Nothing is applied at preview time any more (TASKS.md #148) — this
          // marks "the change card is ready to review", not a write.
          'Preview ready — no workbook changes until Accept',
          {
            changeSetId: meta?.changeSetId,
            changeCount: changes.length,
            actionCount: actions.length,
          },
        );
      } catch (error) {
        frontendTelemetry.logPreviewFail(error, actions);
        console.error('[Cellix] Preview apply failed; use Accept to apply changes:', error);
      }
    },
    [],
  );

  const clearActionPreview = useCallback(async () => {
    if (previewManager.active) {
      await previewManager.reject();
    }
  }, []);

  useEffect(() => {
    frontendTelemetry.installConsoleCapture();
  }, []);

  useEffect(() => {
    void resolveWorkbookKey().then(setWorkbookKey);
  }, []);

  useEffect(() => {
    void resolveWorkbookId().then(setWorkbookId);
  }, []);

  useEffect(() => {
    if (workbookKey) {
      frontendTelemetry.setContext({ workbookKey });
    }
  }, [workbookKey]);
  useEffect(() => {
    if (!workbookKey) return;
    const stored = loadChatSessions(workbookKey);
    if (stored?.assistantMode && isAssistantMode(stored.assistantMode)) {
      setMode(stored.assistantMode);
    }
  }, [workbookKey]);

  const {
    sessions,
    activeSessionId,
    turns,
    activeTurnId,
    conversationId,
    isWaitingForResponse,
    isWaitingClarification,
    activeClarification,
    sendMessage,
    answerQuestion,
    answerClarification,
    dismissClarification,
    acceptActions,
    acceptAllActions,
    rejectActions,
    endConversation,
    newChat,
    selectSession,
    closeSession,
    toggleThinking,
    markAnswerComplete,
  } = useConversation({
    workbookKey,
    workbookId,
    onActions: applyActionsWithAudit,
    onPreviewActions: previewActions,
    onClearPreview: clearActionPreview,
    autoApplyActions: !previewEnabled,
    previewEnabled,
    isChangeSetApplied,
  });

  useEffect(() => {
    frontendTelemetry.setContext({ conversationId });
  }, [conversationId]);

  const handleAcceptActions = useCallback(
    async (turnId: string, blockId: string) => {
      if (applyInProgressRef.current || isApplying) return;
      applyInProgressRef.current = true;
      setIsApplying(true);
      try {
        const turn = turns.find((t) => t.id === turnId);
        const block = turn?.blocks.find((b) => b.id === blockId && b.type === 'actions');
        if (block && block.type === 'actions') {
          frontendTelemetry.logAcceptClick(block.actions, {
            changeSetId: block.changeSetId,
            source: 'actionCard',
          });
        }
        await acceptActions(turnId, blockId);
      } finally {
        applyInProgressRef.current = false;
        setIsApplying(false);
      }
    },
    [acceptActions, isApplying, turns],
  );

  /**
   * Accept this step and every remaining one — TASKS.md #160.
   *
   * Shares the same in-flight guard as single Accept, so the two can never run
   * concurrently against the workbook.
   */
  const handleAcceptAllActions = useCallback(
    async (turnId: string, fromBlockId: string) => {
      if (applyInProgressRef.current || isApplying) return;
      applyInProgressRef.current = true;
      setIsApplying(true);
      try {
        const turn = turns.find((t) => t.id === turnId);
        const block = turn?.blocks.find((b) => b.id === fromBlockId && b.type === 'actions');
        if (block && block.type === 'actions') {
          frontendTelemetry.logAcceptClick(block.actions, {
            changeSetId: block.changeSetId,
            source: 'actionCardAcceptAll',
          });
        }
        await acceptAllActions(turnId, fromBlockId);
      } finally {
        applyInProgressRef.current = false;
        setIsApplying(false);
      }
    },
    [acceptAllActions, isApplying, turns],
  );

  const handleRejectActions = useCallback(
    async (turnId: string, blockId: string) => {
      const turn = turns.find((t) => t.id === turnId);
      const block = turn?.blocks.find((b) => b.id === blockId && b.type === 'actions');
      frontendTelemetry.logReject({
        changeSetId: block && block.type === 'actions' ? block.changeSetId : undefined,
        source: 'actionCard',
      });
      await rejectActions(turnId, blockId);
    },
    [rejectActions, turns],
  );

  const readWorkbookData = useCallback(async () => {
    setIsReadingWorkbook(true);
    try {
      // Spec 09 item 1: reuse selection-time prebuild when still fresh.
      const resolved = await getContextForSend();
      return {
        sheetData: resolved.sheetData,
        workbookContext: resolved.workbookContext,
        promptContext: resolved.promptContext,
      };
    } catch (err) {
      console.error('[Cellix] Workbook read failed:', err);
      return {
        sheetData: [] as unknown[][],
        workbookContext: {
          sheets: [],
          activeSheet: 'Sheet1',
        },
        promptContext: undefined,
      };
    } finally {
      setIsReadingWorkbook(false);
    }
  }, []);

  const handleSend = async (message: string, modeOverride?: AssistantMode) => {
    if (!message.trim() || isWaitingForResponse) return;

    const effectiveMode = modeOverride ?? mode;

    if (modeOverride && modeOverride !== mode) {
      handleModeChange(modeOverride);
    }

    const { sheetData, workbookContext, promptContext } = await readWorkbookData();
    await sendMessage(message.trim(), sheetData, workbookContext, promptContext, {
      mode: effectiveMode,
    });
  };

  /**
   * Regenerate (same text) or edit-and-resend (new text) an existing turn
   * in place, rather than appending a duplicate further down the thread —
   * the message stays anchored, only its answer changes.
   */
  const handleRegenerate = async (turnId: string, overrideMessage?: string) => {
    if (isWaitingForResponse) return;
    const target = turns.find((t) => t.id === turnId);
    const text = (overrideMessage ?? target?.userMessage ?? '').trim();
    if (!text) return;

    const { sheetData, workbookContext, promptContext } = await readWorkbookData();
    await sendMessage(text, sheetData, workbookContext, promptContext, {
      mode,
      regenerateTurnId: turnId,
    });
  };

  const handleRunAsAction = useCallback(
    (message: string) => {
      setMode('action');
      if (workbookKey) {
        const existing = loadChatSessions(workbookKey);
        saveChatSessions(workbookKey, {
          activeSessionId: existing?.activeSessionId ?? null,
          sessions: existing?.sessions ?? [],
          assistantMode: 'action',
        });
      }
      void handleSend(message, 'action');
    },
    // handleSend is defined inline each render; intentionally omitted from deps.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [mode, isWaitingForResponse],
  );

  const handleAnswerQuestion = async (answer: string) => {
    if (isWaitingForResponse) return;
    const { sheetData, workbookContext, promptContext } = await readWorkbookData();
    await answerQuestion(answer, sheetData, workbookContext, promptContext, { mode });
  };

  const handleClarificationAnswer = async (answer: string) => {
    const { sheetData, workbookContext, promptContext } = await readWorkbookData();
    await answerClarification(answer, sheetData, workbookContext, promptContext, { mode });
  };

  const handleRevertHistoryEntry = useCallback(
    async (_changeSetId: string, inverseActions: SheetAction[]) => {
      if (inverseActions.length > 0) {
        await ActionEngine.applyActions(inverseActions);
      }
    },
    [],
  );

  const handleRestoreCheckpoint = useCallback(
    async (result: RestoreResult) => {
      // TASKS.md #29/#31 — AD-1: the backend never writes to the live workbook.
      // restoreCheckpoint() only computed and verified the inverse actions;
      // applying them for real happens here, the same Office.js write path
      // every other accept/revert goes through.
      if (result.inverseActions.length > 0) {
        await ActionEngine.applyActions(result.inverseActions);
      }
    },
    [],
  );

  return (
    <ConversationPanel
      sessions={sessions}
      activeSessionId={activeSessionId}
      turns={turns}
      activeTurnId={activeTurnId}
      conversationId={conversationId}
      isWaitingForResponse={isWaitingForResponse || isReadingWorkbook}
      isWaitingClarification={isWaitingClarification}
      activeClarification={activeClarification}
      previewEnabled={previewEnabled}
      mode={mode}
      onModeChange={handleModeChange}
      onRunAsAction={handleRunAsAction}
      compareResult={compareResult}
      isComparing={isComparing}
      onCloseCompare={() => setCompareResult(null)}
      onSend={handleSend}
      onStop={endConversation}
      onNewChat={newChat}
      onSelectSession={selectSession}
      onCloseSession={closeSession}
      onAcceptActions={handleAcceptActions}
      onAcceptAllActions={handleAcceptAllActions}
      onRejectActions={handleRejectActions}
      onAnswerQuestion={handleAnswerQuestion}
      onClarificationAnswer={handleClarificationAnswer}
      onClarificationDismiss={dismissClarification}
      onToggleThinking={toggleThinking}
      onAnswerComplete={markAnswerComplete}
      onFollowUp={handleSend}
      onRegenerate={handleRegenerate}
      onRevertHistoryEntry={handleRevertHistoryEntry}
      workbookId={workbookId}
      onRestoreCheckpoint={handleRestoreCheckpoint}
      isApplyingActions={isApplying}
    />
  );
};

export default App;
