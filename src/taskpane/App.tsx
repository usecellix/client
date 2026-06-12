import React, { useCallback, useMemo, useState } from 'react';
import ConversationPanel from '@/components/ConversationPanel/ConversationPanel';
import { CompareResult } from '@/components/SheetCompareView/SheetCompareView';
import { useConversation } from '@/hooks/useConversation';
import { ActionEngine } from '@/utils/actionEngine';
import { previewManager, DiffItem } from '@/services/previewManager';
import { buildWorkbookContext, compareSheets } from '@/services/sheetContextBuilder';
import { SheetAction } from '@/types/sheet-actions';
import '@/styles/conversation-panel.css';
import './taskpane.css';

/* global Excel */

const PREVIEW_KEY = 'cellix.previewEnabled';

function readPreviewEnabled(): boolean {
  const stored = localStorage.getItem(PREVIEW_KEY);
  return stored === null ? true : stored === 'true';
}

const App: React.FC = () => {
  const [previewEnabled, setPreviewEnabled] = useState(readPreviewEnabled);
  const [diffItems, setDiffItems] = useState<DiffItem[]>([]);
  const [diffSummary, setDiffSummary] = useState('');
  const [hasPendingPreview, setHasPendingPreview] = useState(false);
  const [isApplying, setIsApplying] = useState(false);
  const [selectedSheets, setSelectedSheets] = useState<string[]>([]);
  const [compareResult, setCompareResult] = useState<CompareResult | null>(null);
  const [isComparing, setIsComparing] = useState(false);

  const togglePreview = useCallback(() => {
    setPreviewEnabled((prev) => {
      const next = !prev;
      localStorage.setItem(PREVIEW_KEY, String(next));
      return next;
    });
  }, []);

  const clearPreviewState = useCallback(() => {
    setHasPendingPreview(false);
    setDiffItems([]);
    setDiffSummary('');
  }, []);

  const applyActions = useCallback(async (actions: SheetAction[], explanation: string) => {
    if (!actions.length) return;

    if (previewManager.active) {
      await previewManager.accept();
      clearPreviewState();
    } else {
      await ActionEngine.applyActions(actions);
    }

    console.info('[Cellix] Applied actions:', explanation, actions);
  }, [clearPreviewState]);

  const previewActions = useCallback(async (actions: SheetAction[], explanation: string) => {
    if (!actions.length) return;

    const items = await previewManager.render({
      actions,
      summary: explanation,
    });
    setDiffItems(items);
    setDiffSummary(explanation);
    setHasPendingPreview(true);
  }, []);

  const clearActionPreview = useCallback(async () => {
    if (previewManager.active) {
      await previewManager.reject();
    }
    clearPreviewState();
  }, [clearPreviewState]);

  const {
    turns,
    activeTurnId,
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
    onActions: applyActions,
    onPreviewActions: previewActions,
    onClearPreview: clearActionPreview,
    autoApplyActions: !previewEnabled,
    previewEnabled,
  });

  const findPendingActionBlock = useCallback(() => {
    for (const turn of turns) {
      const block = turn.blocks.find(
        (b) => b.type === 'actions' && b.proposalStatus === 'pending',
      );
      if (block && block.type === 'actions') {
        return { turnId: turn.id, blockId: block.id };
      }
    }
    return null;
  }, [turns]);

  const handlePreviewAccept = useCallback(async () => {
    const pending = findPendingActionBlock();
    if (!pending && !previewManager.active) return;

    setIsApplying(true);
    try {
      if (pending) {
        await acceptActions(pending.turnId, pending.blockId);
      } else {
        await previewManager.accept();
      }
      clearPreviewState();
    } catch (error) {
      console.error('[Cellix] Failed to apply previewed changes:', error);
    } finally {
      setIsApplying(false);
    }
  }, [acceptActions, clearPreviewState, findPendingActionBlock]);

  const handlePreviewReject = useCallback(async () => {
    const pending = findPendingActionBlock();
    if (pending) {
      await rejectActions(pending.turnId, pending.blockId);
    } else {
      await clearActionPreview();
    }
    clearPreviewState();
  }, [clearActionPreview, clearPreviewState, findPendingActionBlock, rejectActions]);

  const pendingPreview = useMemo(() => {
    if (!previewEnabled || !hasPendingPreview) return null;
    return {
      items: diffItems,
      summary: diffSummary,
      isApplying,
      onAccept: handlePreviewAccept,
      onReject: handlePreviewReject,
    };
  }, [
    previewEnabled,
    hasPendingPreview,
    diffItems,
    diffSummary,
    isApplying,
    handlePreviewAccept,
    handlePreviewReject,
  ]);

  const readWorkbookData = useCallback(async () => {
    try {
      const { context, activeSheetData } = await buildWorkbookContext(selectedSheets);
      return { sheetData: activeSheetData, workbookContext: context };
    } catch (err) {
      console.error('Error reading workbook:', err);
      return {
        sheetData: [] as unknown[][],
        workbookContext: {
          sheets: [],
          activeSheet: 'Sheet1',
        },
      };
    }
  }, [selectedSheets]);

  const handleSheetSelectionChange = useCallback((sheets: string[]) => {
    setSelectedSheets(sheets);
  }, []);

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

  const handleSend = async (message: string) => {
    if (!message.trim() || isWaitingForResponse) return;
    const { sheetData, workbookContext } = await readWorkbookData();
    await sendMessage(message.trim(), sheetData, workbookContext);
  };

  const handleAnswerQuestion = async (answer: string) => {
    if (isWaitingForResponse) return;
    const { sheetData, workbookContext } = await readWorkbookData();
    await answerQuestion(answer, sheetData, workbookContext);
  };

  const handleClarificationAnswer = async (answer: string) => {
    const { sheetData, workbookContext } = await readWorkbookData();
    await answerClarification(answer, sheetData, workbookContext);
  };

  return (
    <ConversationPanel
      turns={turns}
      activeTurnId={activeTurnId}
      isWaitingForResponse={isWaitingForResponse}
      isWaitingClarification={isWaitingClarification}
      activeClarification={activeClarification}
      previewEnabled={previewEnabled}
      onTogglePreview={togglePreview}
      onSheetSelectionChange={handleSheetSelectionChange}
      onCompareSheets={handleCompare}
      compareResult={compareResult}
      isComparing={isComparing}
      onCloseCompare={() => setCompareResult(null)}
      onSend={handleSend}
      onStop={endConversation}
      onClear={clearConversation}
      onAcceptActions={acceptActions}
      onRejectActions={rejectActions}
      onAnswerQuestion={handleAnswerQuestion}
      onClarificationAnswer={handleClarificationAnswer}
      onClarificationDismiss={dismissClarification}
      onToggleThinking={toggleThinking}
      onAnswerComplete={markAnswerComplete}
      onFollowUp={handleSend}
      pendingPreview={pendingPreview}
    />
  );
};

export default App;
