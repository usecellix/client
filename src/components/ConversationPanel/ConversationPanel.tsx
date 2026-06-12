import React, { useRef, useEffect, KeyboardEvent } from 'react';
import { ArrowRight, Square } from 'lucide-react';
import { ClarificationPanel } from '@/components/ClarificationPanel/ClarificationPanel';
import { SheetSelector } from '@/components/SheetSelector/SheetSelector';
import { SheetCompareView, CompareResult } from '@/components/SheetCompareView/SheetCompareView';
import { ClarificationPayload } from '@/types/cellix.types';
import { PreviewSummaryBar } from '@/components/PreviewSummaryBar/PreviewSummaryBar';
import { DiffItem } from '@/services/previewManager';
import TurnRenderer from './TurnRenderer';

interface PanelInputProps {
  onSend: (message: string) => void;
  onStop?: () => void;
  disabled?: boolean;
  isProcessing?: boolean;
  isWaitingClarification?: boolean;
  placeholder?: string;
}

const SUGGESTIONS = [
  "What's in cell A1?",
  'Calculate the total',
  'Explain this sheet to me',
];

interface EmptyStateProps {
  onSuggestion: (text: string) => void;
}

export const EmptyState: React.FC<EmptyStateProps> = ({ onSuggestion }) => (
  <div className="cellix-empty">
    <div className="cellix-brand">
      <div className="cellix-logo-icon">C</div>
      <span className="cellix-brand-text">CELLIX</span>
    </div>
    <div className="cellix-version">AI assistant for your active worksheet</div>
    {SUGGESTIONS.map((suggestion) => (
      <button
        key={suggestion}
        type="button"
        className="cellix-suggestion"
        onClick={() => onSuggestion(suggestion)}
      >
        {suggestion}
        <span>→</span>
      </button>
    ))}
  </div>
);

export const PanelInput: React.FC<PanelInputProps> = ({
  onSend,
  onStop,
  disabled = false,
  isProcessing = false,
  isWaitingClarification = false,
  placeholder = 'Ask anything about your spreadsheet…',
}) => {
  const [message, setMessage] = React.useState('');
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const handleSend = () => {
    if (message.trim() && !disabled) {
      onSend(message.trim());
      setMessage('');
      if (textareaRef.current) textareaRef.current.style.height = 'auto';
    }
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const inputDisabled = disabled || isWaitingClarification;
  const resolvedPlaceholder = isWaitingClarification
    ? '⏸ Answer the question above first…'
    : isProcessing
      ? 'Processing your request…'
      : placeholder;

  return (
    <div className="cellix-input-area">
      <div
        className={`cellix-input-shell ${isProcessing ? 'processing' : ''} ${
          isWaitingClarification ? 'clarifying' : ''
        }`}
      >
        <textarea
          ref={textareaRef}
          value={message}
          onChange={(e) => {
            setMessage(e.target.value);
            e.target.style.height = 'auto';
            e.target.style.height = `${Math.min(e.target.scrollHeight, 120)}px`;
          }}
          onKeyDown={handleKeyDown}
          placeholder={resolvedPlaceholder}
          disabled={inputDisabled}
          rows={1}
        />
        <div className="cellix-input-toolbar">
          <div style={{ fontSize: 11, color: 'var(--cx-gray-500)' }}>
            {isWaitingClarification
              ? 'Clarification required'
              : isProcessing
                ? 'Action mode'
                : 'Enter to send'}
          </div>
          {isProcessing ? (
            <button type="button" className="cellix-stop-btn" onClick={onStop} aria-label="Stop">
              <Square size={10} fill="white" color="white" />
            </button>
          ) : (
            <button
              type="button"
              className="cellix-send-btn"
              onClick={handleSend}
              disabled={!message.trim() || inputDisabled}
              aria-label="Send"
            >
              <ArrowRight size={14} />
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

interface ConversationPanelProps {
  turns: import('@/types/conversationTurn').ConversationTurn[];
  activeTurnId: string | null;
  isWaitingForResponse: boolean;
  isWaitingClarification: boolean;
  activeClarification: ClarificationPayload | null;
  previewEnabled: boolean;
  onTogglePreview: () => void;
  onSheetSelectionChange: (sheets: string[]) => void;
  onCompareSheets: (sheetA: string, sheetB: string) => void;
  compareResult: CompareResult | null;
  isComparing: boolean;
  onCloseCompare: () => void;
  onSend: (message: string) => void;
  onStop: () => void;
  onClear: () => void;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onAnswerQuestion: (answer: string) => void;
  onClarificationAnswer: (answer: string) => void;
  onClarificationDismiss: () => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onFollowUp: (text: string) => void;
  pendingPreview?: {
    items: DiffItem[];
    summary: string;
    isApplying: boolean;
    onAccept: () => void;
    onReject: () => void;
  } | null;
}

const ConversationPanel: React.FC<ConversationPanelProps> = ({
  turns,
  activeTurnId,
  isWaitingForResponse,
  isWaitingClarification,
  activeClarification,
  previewEnabled,
  onTogglePreview,
  onSheetSelectionChange,
  onCompareSheets,
  compareResult,
  isComparing,
  onCloseCompare,
  onSend,
  onStop,
  onClear,
  onAcceptActions,
  onRejectActions,
  onAnswerQuestion,
  onClarificationAnswer,
  onClarificationDismiss,
  onToggleThinking,
  onAnswerComplete,
  onFollowUp,
  pendingPreview = null,
}) => {
  const contentRef = useRef<HTMLDivElement>(null);
  const activeTurn = turns.find((t) => t.id === activeTurnId);

  useEffect(() => {
    contentRef.current?.scrollTo({ top: contentRef.current.scrollHeight, behavior: 'smooth' });
  }, [turns, isWaitingForResponse, activeClarification]);

  return (
    <div className="cellix-panel">
      <div className="cellix-topbar">
        <span className="cellix-topbar-title">Cellix</span>
        <div className="cellix-topbar-actions">
          <button
            type="button"
            onClick={onTogglePreview}
            className={`cellix-preview-toggle ${previewEnabled ? 'on' : 'off'}`}
            title="Toggle preview before applying changes"
          >
            <span>{previewEnabled ? '👁' : '👁‍🗨'}</span>
            Preview {previewEnabled ? 'ON' : 'OFF'}
          </button>
          {turns.length > 0 && (
            <button type="button" className="cellix-icon-btn" onClick={onClear} title="New chat">
              +
            </button>
          )}
        </div>
      </div>

      {activeTurn && (
        <div className="cellix-tabbar-wrap">
          <div className="cellix-tab">
            {isWaitingForResponse && activeTurn.id === activeTurnId ? (
              <span className="cellix-spinner" />
            ) : (
              <span
                className="cellix-status-dot"
                style={{ width: 7, height: 7 }}
              />
            )}
            <span className="cellix-tab-label">{activeTurn.tabLabel}</span>
          </div>
        </div>
      )}

      <div className="cellix-content" ref={contentRef}>
        {turns.length === 0 ? (
          <EmptyState onSuggestion={onSend} />
        ) : (
          turns.map((turn) => (
            <TurnRenderer
              key={turn.id}
              turn={turn}
              isActive={turn.id === activeTurnId}
              isWaiting={isWaitingForResponse}
              previewEnabled={previewEnabled}
              onAcceptActions={onAcceptActions}
              onRejectActions={onRejectActions}
              onAnswerQuestion={onAnswerQuestion}
              onToggleThinking={onToggleThinking}
              onAnswerComplete={onAnswerComplete}
              onFollowUp={onFollowUp}
            />
          ))
        )}
      </div>

      <div className="cellix-bottom-tools">
        <SheetSelector
          onSelectionChange={onSheetSelectionChange}
          onCompare={onCompareSheets}
        />
        {(isComparing || compareResult) && (
          <SheetCompareView
            result={compareResult}
            isLoading={isComparing}
            onClose={onCloseCompare}
          />
        )}
        {activeClarification && (
          <ClarificationPanel
            payload={activeClarification}
            onAnswer={onClarificationAnswer}
            onDismiss={onClarificationDismiss}
          />
        )}
      </div>

      {pendingPreview && (
        <PreviewSummaryBar
          items={pendingPreview.items}
          summary={pendingPreview.summary}
          onAccept={pendingPreview.onAccept}
          onReject={pendingPreview.onReject}
          isApplying={pendingPreview.isApplying}
        />
      )}

      <PanelInput
        onSend={onSend}
        onStop={onStop}
        disabled={isWaitingForResponse}
        isProcessing={isWaitingForResponse}
        isWaitingClarification={isWaitingClarification}
      />
    </div>
  );
};

export default ConversationPanel;
