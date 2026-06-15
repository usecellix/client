import React, { useCallback, useMemo, useRef, useState } from 'react';
import ConversationPanel from '@/components/ConversationPanel/ConversationPanel';
import { CompareResult } from '@/components/SheetCompareView/SheetCompareView';
import { useConversation, PreviewActionsMeta } from '@/hooks/useConversation';
import { ActionEngine } from '@/utils/actionEngine';
import { CellChange } from '@/types/changeSet';
import { previewManager } from '@/services/previewManager';
import { markChangeSetApplied } from '@/services/auditService';
import { buildWorkbookContext, compareSheets } from '@/services/sheetContextBuilder';
import { SheetAction } from '@/types/sheet-actions';
import { AssistantMode, DEFAULT_ASSISTANT_MODE, isAssistantMode } from '@/types/mode';
import '@/styles/conversation-panel.css';
import './taskpane.css';

/* global Excel */

const PREVIEW_KEY = 'cellix.previewEnabled';
const MODE_KEY = 'cellix.assistantMode';

function readPreviewEnabled(): boolean {
  const stored = localStorage.getItem(PREVIEW_KEY);
  return stored === null ? true : stored === 'true';
}

function readMode(): AssistantMode {
  const stored = localStorage.getItem(MODE_KEY);
  return isAssistantMode(stored) ? stored : DEFAULT_ASSISTANT_MODE;
}

const App: React.FC = () => {
  const [previewEnabled, setPreviewEnabled] = useState(readPreviewEnabled);
  const [mode, setMode] = useState<AssistantMode>(readMode);
  const [serverChanges, setServerChanges] = useState<CellChange[]>([]);
  const [pendingChangeSetId, setPendingChangeSetId] = useState<string | undefined>();
  const [diffSummary, setDiffSummary] = useState('');
  const [hasPendingPreview, setHasPendingPreview] = useState(false);
  const [isApplying, setIsApplying] = useState(false);
  const [compareResult, setCompareResult] = useState<CompareResult | null>(null);
  const [isComparing, setIsComparing] = useState(false);
  const [isReadingWorkbook, setIsReadingWorkbook] = useState(false);
  const [refinementChangeSetId, setRefinementChangeSetId] = useState<string | null>(null);
  const [quickEditMode, setQuickEditMode] = useState(false);
  const applyInProgressRef = useRef(false);
  const appliedChangeSetIdsRef = useRef<Set<string>>(new Set());

  const isChangeSetApplied = useCallback((changeSetId?: string) => {
    return Boolean(changeSetId && appliedChangeSetIdsRef.current.has(changeSetId));
  }, []);

  const togglePreview = useCallback(() => {
    setPreviewEnabled((prev) => {
      const next = !prev;
      localStorage.setItem(PREVIEW_KEY, String(next));
      return next;
    });
  }, []);

  const handleModeChange = useCallback((next: AssistantMode) => {
    setMode(next);
    localStorage.setItem(MODE_KEY, next);
  }, []);

  const clearPreviewState = useCallback(() => {
    setHasPendingPreview(false);
    setServerChanges([]);
    setPendingChangeSetId(undefined);
    setDiffSummary('');
  }, []);

  const applyActionsWithAudit = useCallback(
    async (actions: SheetAction[], explanation: string, meta?: PreviewActionsMeta) => {
      if (!actions.length) return;

      if (previewManager.active) {
        await previewManager.accept();
      } else {
        await ActionEngine.applyActions(actions);
      }

      clearPreviewState();

      if (meta?.changeSetId) {
        appliedChangeSetIdsRef.current.add(meta.changeSetId);
        try {
          await markChangeSetApplied(meta.changeSetId);
          setRefinementChangeSetId(meta.changeSetId);
        } catch (error) {
          console.warn('[Cellix] Failed to mark change set applied:', error);
        }
      }

      console.info('[Cellix] Applied actions:', explanation, actions);
    },
    [clearPreviewState],
  );

  const previewActions = useCallback(
    async (actions: SheetAction[], explanation: string, meta?: PreviewActionsMeta) => {
      if (!actions.length) return;

      const changes = meta?.changes ?? [];
      if (changes.length > 0) {
        await previewManager.highlightChanges(changes, actions);
      } else {
        await previewManager.render({ actions, summary: explanation });
      }

      setServerChanges(changes);
      setPendingChangeSetId(meta?.changeSetId);
      setDiffSummary(explanation);
      setHasPendingPreview(true);
    },
    [],
  );

  const clearActionPreview = useCallback(async () => {
    if (previewManager.active) {
      await previewManager.reject();
    }
    clearPreviewState();
  }, [clearPreviewState]);

  const {
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
    clearConversation,
    toggleThinking,
    markAnswerComplete,
  } = useConversation({
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
      if (pending) {
        await acceptActions(pending.turnId, pending.blockId);
        applied = true;
      } else if (previewManager.active) {
        await previewManager.accept();
        if (pendingChangeSetId) {
          appliedChangeSetIdsRef.current.add(pendingChangeSetId);
          await markChangeSetApplied(pendingChangeSetId);
          setRefinementChangeSetId(pendingChangeSetId);
        }
        applied = true;
      }
    } catch (error) {
      console.error('[Cellix] Failed to apply previewed changes:', error);
    } finally {
      if (applied) {
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
        await acceptActions(turnId, blockId);
      } finally {
        applyInProgressRef.current = false;
        setIsApplying(false);
      }
    },
    [acceptActions, isApplying],
  );

  const handlePreviewReject = useCallback(async () => {
    if (applyInProgressRef.current || isApplying) return;

    applyInProgressRef.current = true;
    setIsApplying(true);
    try {
      const pending = findPendingActionBlock();
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
  }, [clearActionPreview, clearPreviewState, findPendingActionBlock, isApplying, rejectActions]);

  const handleRevertChangeSet = useCallback(
    async (_changeSetId: string, inverseActions: SheetAction[]) => {
      if (!inverseActions.length) return;
      await ActionEngine.applyActions(inverseActions);
    },
    [],
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
      // Always read the entire workbook (all visible sheets) so every request
      // has full cross-sheet context. Sheet selection narrows display/compare
      // only, never the search/operation scope.
      const { context, activeSheetData, promptContext } = await buildWorkbookContext([]);
      return {
        sheetData: activeSheetData,
        workbookContext: context,
        promptContext,
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

  // Sheet selection is display/compare-only; search & operations always use
  // the full workbook context, so no scope state is tracked here.
  const handleSheetSelectionChange = useCallback(() => {}, []);

  const handleCompare = useCallback(async (sheetA: string, sheetB: string) => {
    setIsComparing(true);
    setCompareResult(null);
    try {
      const result = await compareSheets(sheetA, sheetB);
      setCompareResult(result);
    } catch (error) {
      console.error('[Cellix] Sheet compare failed:', error);
    } finally {
      setIsComparing(false);
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

  return (
    <ConversationPanel
      turns={turns}
      activeTurnId={activeTurnId}
      conversationId={conversationId}
      isWaitingForResponse={isWaitingForResponse || isReadingWorkbook}
      isWaitingClarification={isWaitingClarification}
      activeClarification={activeClarification}
      previewEnabled={previewEnabled}
      onTogglePreview={togglePreview}
      mode={mode}
      onModeChange={handleModeChange}
      onRunAsAction={handleRunAsAction}
      onSheetSelectionChange={handleSheetSelectionChange}
      onCompareSheets={handleCompare}
      compareResult={compareResult}
      isComparing={isComparing}
      onCloseCompare={() => setCompareResult(null)}
      onSend={handleSend}
      onStop={endConversation}
      onClear={clearConversation}
      onAcceptActions={handleAcceptActions}
      onRejectActions={rejectActions}
      onAnswerQuestion={handleAnswerQuestion}
      onClarificationAnswer={handleClarificationAnswer}
      onClarificationDismiss={dismissClarification}
      onToggleThinking={toggleThinking}
      onAnswerComplete={markAnswerComplete}
      onFollowUp={handleSend}
      isApplyingActions={isApplying}
      pendingPreview={pendingPreview}
      onRevertChangeSet={handleRevertChangeSet}
      refinementChangeSetId={refinementChangeSetId}
      quickEditMode={quickEditMode}
      onStartQuickEdit={() => setQuickEditMode(true)}
      onCancelQuickEdit={() => setQuickEditMode(false)}
    />
  );
};

export default App;
