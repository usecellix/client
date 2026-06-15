import React, { useRef, useEffect, KeyboardEvent } from 'react';
import { ArrowRight, Square } from 'lucide-react';
import { AssistantMode, ASSISTANT_MODES, ASSISTANT_MODE_META } from '@/types/mode';
import { ClarificationPanel } from '@/components/ClarificationPanel/ClarificationPanel';
import { SheetCompareView, CompareResult } from '@/components/SheetCompareView/SheetCompareView';
import { ClarificationPayload } from '@/types/cellix.types';
import { DiffPreviewPanel } from '@/components/DiffPreviewPanel/DiffPreviewPanel';
import { CellChange } from '@/types/changeSet';
import { SheetAction } from '@/types/sheet-actions';
import PanelHeader from './PanelHeader';
import TurnRenderer from './TurnRenderer';

interface PanelInputProps {
  onSend: (message: string) => void;
  onStop?: () => void;
  disabled?: boolean;
  isProcessing?: boolean;
  isWaitingClarification?: boolean;
  placeholder?: string;
  mode: AssistantMode;
  onModeChange: (mode: AssistantMode) => void;
}

interface ModeSwitchProps {
  mode: AssistantMode;
  onModeChange: (mode: AssistantMode) => void;
  disabled?: boolean;
}

const ModeSwitch: React.FC<ModeSwitchProps> = ({ mode, onModeChange, disabled = false }) => (
  <div className="cellix-mode-switch" role="tablist" aria-label="Assistant mode">
    {ASSISTANT_MODES.map((m) => {
      const meta = ASSISTANT_MODE_META[m];
      const active = m === mode;
      return (
        <button
          key={m}
          type="button"
          role="tab"
          aria-selected={active}
          title={meta.hint}
          className={`cellix-mode-pill ${active ? 'active' : ''}`}
          onClick={() => !disabled && onModeChange(m)}
          disabled={disabled}
        >
          {meta.label}
        </button>
      );
    })}
  </div>
);

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
  mode,
  onModeChange,
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

  const modeLabel = ASSISTANT_MODE_META[mode].label;

  return (
    <div className="cellix-input-area">
      <div
        className={`cellix-input-shell mode-${mode} ${isProcessing ? 'processing' : ''} ${
          isWaitingClarification ? 'clarifying' : ''
        }`}
      >
        <ModeSwitch mode={mode} onModeChange={onModeChange} disabled={isProcessing} />
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
          <div className="cellix-input-hint">
            {isWaitingClarification
              ? 'Clarification required'
              : isProcessing
                ? `${modeLabel} mode — working…`
                : `${modeLabel} mode · Enter to send`}
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
  conversationId?: string | null;
  isWaitingForResponse: boolean;
  isWaitingClarification: boolean;
  activeClarification: ClarificationPayload | null;
  previewEnabled: boolean;
  onTogglePreview: () => void;
  mode: AssistantMode;
  onModeChange: (mode: AssistantMode) => void;
  onRunAsAction: (message: string) => void;
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
  isApplyingActions?: boolean;
  pendingPreview?: {
    changes: CellChange[];
    changeSetId?: string;
    summary: string;
    isApplying: boolean;
    onAccept: () => void;
    onReject: () => void;
  } | null;
  onRevertChangeSet?: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
  refinementChangeSetId?: string | null;
  quickEditMode?: boolean;
  onStartQuickEdit?: () => void;
  onCancelQuickEdit?: () => void;
}

const ConversationPanel: React.FC<ConversationPanelProps> = ({
  turns,
  activeTurnId,
  conversationId = null,
  isWaitingForResponse,
  isWaitingClarification,
  activeClarification,
  previewEnabled,
  onTogglePreview,
  mode,
  onModeChange,
  onRunAsAction,
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
  isApplyingActions = false,
  pendingPreview = null,
  onRevertChangeSet,
  refinementChangeSetId = null,
  quickEditMode = false,
  onStartQuickEdit,
  onCancelQuickEdit,
}) => {
  const contentRef = useRef<HTMLDivElement>(null);
  const activeTurn = turns.find((t) => t.id === activeTurnId);

  useEffect(() => {
    contentRef.current?.scrollTo({ top: contentRef.current.scrollHeight, behavior: 'smooth' });
  }, [turns, isWaitingForResponse, activeClarification]);

  return (
    <div className="cellix-panel">
      <PanelHeader
        previewEnabled={previewEnabled}
        onTogglePreview={onTogglePreview}
        onClear={onClear}
        hasTurns={turns.length > 0}
        conversationId={conversationId ?? null}
        onRevertChangeSet={onRevertChangeSet}
        onSheetSelectionChange={onSheetSelectionChange}
        onCompareSheets={onCompareSheets}
      />

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
              isApplying={isApplyingActions}
              onAcceptActions={onAcceptActions}
              onRejectActions={onRejectActions}
              onAnswerQuestion={onAnswerQuestion}
              onToggleThinking={onToggleThinking}
              onAnswerComplete={onAnswerComplete}
              onFollowUp={onFollowUp}
              onRunAsAction={onRunAsAction}
            />
          ))
        )}
      </div>

      {(isComparing || compareResult || activeClarification) && (
        <div className="cellix-bottom-tools">
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
      )}

      {pendingPreview && (
        <DiffPreviewPanel
          changes={pendingPreview.changes}
          changeSetId={pendingPreview.changeSetId}
          summary={pendingPreview.summary}
          onAccept={pendingPreview.onAccept}
          onReject={pendingPreview.onReject}
          isApplying={pendingPreview.isApplying}
        />
      )}

      {refinementChangeSetId && !quickEditMode && onStartQuickEdit && (
        <div className="cellix-quick-edit-banner">
          <span>Last change applied — refine without re-reading the workbook.</span>
          <button type="button" className="cellix-btn-secondary" onClick={onStartQuickEdit}>
            Quick edit
          </button>
        </div>
      )}

      {quickEditMode && onCancelQuickEdit && (
        <div className="cellix-quick-edit-banner active">
          <span>Quick edit mode — describe how to adjust the last change.</span>
          <button type="button" className="cellix-btn-secondary" onClick={onCancelQuickEdit}>
            Cancel
          </button>
        </div>
      )}

      <PanelInput
        onSend={onSend}
        onStop={onStop}
        disabled={isWaitingForResponse}
        isProcessing={isWaitingForResponse}
        isWaitingClarification={isWaitingClarification}
        mode={mode}
        onModeChange={onModeChange}
        placeholder={
          quickEditMode
            ? 'Describe how to adjust the last change…'
            : mode === 'ask'
              ? 'Ask anything about your workbook…'
              : mode === 'plan'
                ? 'Describe what you want to plan…'
                : 'Describe the change you want to make…'
        }
      />
    </div>
  );
};

export default ConversationPanel;
