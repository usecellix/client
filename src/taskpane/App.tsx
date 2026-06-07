import React, { useCallback } from 'react';
import ConversationPanel from '@/components/ConversationPanel/ConversationPanel';
import { useConversation } from '@/hooks/useConversation';
import { ActionEngine } from '@/utils/actionEngine';
import { SheetAction } from '@/types/sheet-actions';
import { WorkbookContextPayload } from '@/utils/payloadCompressor';
import '@/styles/conversation-panel.css';
import './taskpane.css';

/* global Excel */

const App: React.FC = () => {
  const applyActions = useCallback(async (actions: SheetAction[], explanation: string) => {
    if (!actions.length) return;
    await ActionEngine.applyActions(actions);
    console.info('[Cellix] Applied actions:', explanation, actions);
  }, []);

  const previewActions = useCallback(async (actions: SheetAction[]) => {
    if (!actions.length) return;
    await ActionEngine.previewActions(actions);
  }, []);

  const clearActionPreview = useCallback(async () => {
    await ActionEngine.clearPreview();
  }, []);

  const {
    turns,
    activeTurnId,
    isWaitingForResponse,
    sendMessage,
    answerQuestion,
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
    autoApplyActions: true,
  });

  const readSheetData = async (): Promise<{ sheetData: unknown[][]; workbookContext: WorkbookContextPayload }> => {
    try {
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const sheets = context.workbook.worksheets;
        worksheet.load('name');
        sheets.load('items/name');
        const usedRange = worksheet.getUsedRange();

        if (usedRange) {
          usedRange.load('values');
        }

        await context.sync();

        const sheetData = usedRange ? ((usedRange.values as unknown[][]) || []) : [];
        const workbookContext: WorkbookContextPayload = {
          activeSheet: worksheet.name,
          sheets: sheets.items.map((s) => s.name),
        };

        return { sheetData, workbookContext };
      });
    } catch (err) {
      console.error('Error reading worksheet:', err);
      return { sheetData: [], workbookContext: { activeSheet: 'Sheet1', sheets: ['Sheet1'] } };
    }
  };

  const handleSend = async (message: string) => {
    if (!message.trim() || isWaitingForResponse) return;
    const { sheetData, workbookContext } = await readSheetData();
    await sendMessage(message.trim(), sheetData, workbookContext);
  };

  const handleAnswerQuestion = async (answer: string) => {
    if (isWaitingForResponse) return;
    const { sheetData, workbookContext } = await readSheetData();
    await answerQuestion(answer, sheetData, workbookContext);
  };

  return (
    <ConversationPanel
      turns={turns}
      activeTurnId={activeTurnId}
      isWaitingForResponse={isWaitingForResponse}
      onSend={handleSend}
      onStop={endConversation}
      onClear={clearConversation}
      onAcceptActions={acceptActions}
      onRejectActions={rejectActions}
      onAnswerQuestion={handleAnswerQuestion}
      onToggleThinking={toggleThinking}
      onAnswerComplete={markAnswerComplete}
      onFollowUp={handleSend}
    />
  );
};

export default App;
