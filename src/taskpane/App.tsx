import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import ConversationPanel from '@/components/ConversationPanel/ConversationPanel';
import { CompareResult } from '@/components/SheetCompareView/SheetCompareView';
import { useConversation, PreviewActionsMeta } from '@/hooks/useConversation';
import { ActionEngine } from '@/utils/actionEngine';
import type { CreatedConditionalFormatId, CreatedChartId } from '@/engine/actionEngine';
import { CellChange } from '@/types/changeSet';
import { previewManager } from '@/services/previewManager';
import { markChangeSetApplied } from '@/services/auditService';
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
  const [serverChanges, setServerChanges] = useState<CellChange[]>([]);
  const [pendingChangeSetId, setPendingChangeSetId] = useState<string | undefined>();
  const [diffSummary, setDiffSummary] = useState('');
  const [hasPendingPreview, setHasPendingPreview] = useState(false);
  const [isApplying, setIsApplying] = useState(false);
  const [compareResult, setCompareResult] = useState<CompareResult | null>(null);
  const [isComparing] = useState(false);
  const [isReadingWorkbook, setIsReadingWorkbook] = useState(false);
  const [refinementChangeSetId, setRefinementChangeSetId] = useState<string | null>(null);
  const [quickEditMode, setQuickEditMode] = useState(false);
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

  const clearPreviewState = useCallback(() => {
    setHasPendingPreview(false);
    setServerChanges([]);
    setPendingChangeSetId(undefined);
    setDiffSummary('');
  }, []);

  const applyActionsWithAudit = useCallback(
    async (actions: SheetAction[], explanation: string, meta?: PreviewActionsMeta) => {
      if (!actions.length) return;

      frontendTelemetry.logAcceptClick(actions, {
        changeSetId: meta?.changeSetId,
        source: 'applyActionsWithAudit',
      });

      let createdConditionalFormatIds: CreatedConditionalFormatId[] | undefined;
      let createdChartIds: CreatedChartId[] | undefined;
      try {
        if (previewManager.active) {
          const result = await previewManager.accept();
          createdConditionalFormatIds = result?.createdConditionalFormatIds;
          createdChartIds = result?.createdChartIds;
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
        }

        clearPreviewState();

        if (meta?.changeSetId) {
          appliedChangeSetIdsRef.current.add(meta.changeSetId);
          try {
            await markChangeSetApplied(meta.changeSetId, createdConditionalFormatIds, createdChartIds);
            setRefinementChangeSetId(meta.changeSetId);
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
    [clearPreviewState],
  );

  const previewActions = useCallback(
    async (actions: SheetAction[], explanation: string, meta?: PreviewActionsMeta) => {
      if (!actions.length) return;

      const changes = meta?.changes ?? [];
      // Show Accept/Reject even if Office.js preview apply fails — Accept retries apply.
      setServerChanges(changes);
      setPendingChangeSetId(meta?.changeSetId);
      setDiffSummary(explanation);
      setHasPendingPreview(true);

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
          'Preview structural apply succeeded',
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
    clearPreviewState();
  }, [clearPreviewState]);

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
    onChangeSetApplied: (changeSetId) => {
      setRefinementChangeSetId(changeSetId);
    },
    autoApplyActions: !previewEnabled,
    previewEnabled,
    isChangeSetApplied,
  });

  useEffect(() => {
    frontendTelemetry.setContext({ conversationId });
  }, [conversationId]);

  const findPendingActionBlock = useCallback(() => {
    for (const turn of turns) {
      const block = turn.blocks.find(
        (b) => b.type === 'actions' && b.proposalStatus === 'pending',
      );
      if (block && block.type === 'actions') {
        return {
          turnId: turn.id,
          blockId: block.id,
          changeSetId: block.changeSetId,
          actions: block.actions,
        };
      }
    }
    return null;
  }, [turns]);

  const handlePreviewAccept = useCallback(async () => {
    if (applyInProgressRef.current || isApplying) return;

    const pending = findPendingActionBlock();
    if (!pending && !previewManager.active) return;

    applyInProgressRef.current = true;
    setIsApplying(true);
    let applied = false;
    try {
      frontendTelemetry.logAcceptClick(pending?.actions ?? [], {
        changeSetId: pending?.changeSetId ?? pendingChangeSetId,
        source: 'previewSummaryBar',
      });
      if (pending) {
        await acceptActions(pending.turnId, pending.blockId);
        applied = true;
      } else if (previewManager.active) {
        const result = await previewManager.accept();
        if (pendingChangeSetId) {
          appliedChangeSetIdsRef.current.add(pendingChangeSetId);
          try {
            await markChangeSetApplied(
              pendingChangeSetId,
              result?.createdConditionalFormatIds,
              result?.createdChartIds,
            );
            setRefinementChangeSetId(pendingChangeSetId);
          } catch (error) {
            // Same pattern as applyActionsWithAudit (spec 22 Bug 3): do not
            // leave a failed apply marked as applied — UI must not show Applied.
            appliedChangeSetIdsRef.current.delete(pendingChangeSetId);
            throw error;
          }
        }
        applied = true;
        frontendTelemetry.logAcceptSuccess([], {
          changeSetId: pendingChangeSetId,
          explanation: 'Preview accept (no pending block)',
        });
      }
    } catch (error) {
      frontendTelemetry.logAcceptFail(error, pending?.actions ?? [], {
        changeSetId: pending?.changeSetId ?? pendingChangeSetId,
      });
      console.error('[Cellix] Failed to apply previewed changes:', error);
    } finally {
      if (applied) {
        markPendingWorkbookContextStale();
        clearPreviewState();
      }
      applyInProgressRef.current = false;
      setIsApplying(false);
    }
  }, [acceptActions, clearPreviewState, findPendingActionBlock, isApplying, pendingChangeSetId]);

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

  const handlePreviewReject = useCallback(async () => {
    if (applyInProgressRef.current || isApplying) return;

    applyInProgressRef.current = true;
    setIsApplying(true);
    try {
      const pending = findPendingActionBlock();
      frontendTelemetry.logReject({
        changeSetId: pending?.changeSetId ?? pendingChangeSetId,
        source: 'previewSummaryBar',
      });
      if (pending) {
        await rejectActions(pending.turnId, pending.blockId);
      } else {
        await clearActionPreview();
      }
      clearPreviewState();
    } catch (error) {
      console.error('[Cellix] Failed to reject previewed changes:', error);
    } finally {
      applyInProgressRef.current = false;
      setIsApplying(false);
    }
  }, [
    clearActionPreview,
    clearPreviewState,
    findPendingActionBlock,
    isApplying,
    pendingChangeSetId,
    rejectActions,
  ]);

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

  const pendingPreview = useMemo(() => {
    if (!previewEnabled || !hasPendingPreview) return null;
    return {
      changes: serverChanges,
      changeSetId: pendingChangeSetId,
      summary: diffSummary,
      isApplying,
      onAccept: handlePreviewAccept,
      onReject: handlePreviewReject,
    };
  }, [
    previewEnabled,
    hasPendingPreview,
    serverChanges,
    pendingChangeSetId,
    diffSummary,
    isApplying,
    handlePreviewAccept,
    handlePreviewReject,
  ]);

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

    if (quickEditMode && refinementChangeSetId) {
      await sendMessage(message.trim(), [[]], undefined, undefined, {
        refinementChangeSetId,
        mode: 'action',
      });
      setQuickEditMode(false);
      return;
    }

    const { sheetData, workbookContext, promptContext } = await readWorkbookData();
    await sendMessage(message.trim(), sheetData, workbookContext, promptContext, {
      mode: effectiveMode,
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
    [mode, quickEditMode, refinementChangeSetId, isWaitingForResponse],
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
      clearPreviewState();
      setQuickEditMode(false);
      setRefinementChangeSetId(null);
    },
    [clearPreviewState],
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
      clearPreviewState();
      setQuickEditMode(false);
      setRefinementChangeSetId(null);
    },
    [clearPreviewState],
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
      onRejectActions={handleRejectActions}
      onAnswerQuestion={handleAnswerQuestion}
      onClarificationAnswer={handleClarificationAnswer}
      onClarificationDismiss={dismissClarification}
      onToggleThinking={toggleThinking}
      onAnswerComplete={markAnswerComplete}
      onFollowUp={handleSend}
      onRevertHistoryEntry={handleRevertHistoryEntry}
      workbookId={workbookId}
      onRestoreCheckpoint={handleRestoreCheckpoint}
      isApplyingActions={isApplying}
      pendingPreview={pendingPreview}
      refinementChangeSetId={refinementChangeSetId}
      quickEditMode={quickEditMode}
      onStartQuickEdit={() => setQuickEditMode(true)}
      onCancelQuickEdit={() => setQuickEditMode(false)}
    />
  );
};

export default App;
