import React, { useRef, useEffect, useState, KeyboardEvent } from 'react';
import {
  ArrowRight,
  AtSign,
  Check,
  ChevronDown,
  FileText,
  Folder,
  Grid3X3,
  HelpCircle,
  ListChecks,
  Paperclip,
  PencilLine,
  Square,
  X,
} from 'lucide-react';
import { AssistantMode, ASSISTANT_MODES, ASSISTANT_MODE_META } from '@/types/mode';
import { SheetCompareView, CompareResult } from '@/components/SheetCompareView/SheetCompareView';
import { ClarificationPayload } from '@/types/cellix.types';
import { PreviewSummaryBar } from '@/components/PreviewSummaryBar/PreviewSummaryBar';
import { isTurnPresentationComplete } from '@/utils/turnPresentation';
import { ChangeHistoryPanel } from '@/components/ChangeHistoryPanel/ChangeHistoryPanel';
import { CellChange, formatCellValue } from '@/types/changeSet';
import { DiffItem } from '@/services/previewManager';
import { SheetAction } from '@/types/sheet-actions';
import PanelHeader from './PanelHeader';
import TurnRenderer from './TurnRenderer';

/* global Excel */

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

interface ReferenceOption {
  id: string;
  label: string;
  detail: string;
  insertText: string;
  type: 'sheet' | 'range' | 'group';
}

type ComposerSegment =
  | {
      id: string;
      type: 'text';
      text: string;
    }
  | {
      id: string;
      type: 'reference';
      reference: ReferenceOption;
    };

const modeIcon = {
  ask: HelpCircle,
  action: PencilLine,
  plan: ListChecks,
} satisfies Record<AssistantMode, typeof HelpCircle>;

const ModeSwitch: React.FC<ModeSwitchProps> = ({ mode, onModeChange, disabled = false }) => {
  const [open, setOpen] = React.useState(false);
  const wrapRef = useRef<HTMLDivElement>(null);
  const activeMeta = ASSISTANT_MODE_META[mode];
  const ActiveIcon = modeIcon[mode];

  useEffect(() => {
    if (!open) return;

    const handlePointerDown = (event: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(event.target as Node)) {
        setOpen(false);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    return () => document.removeEventListener('mousedown', handlePointerDown);
  }, [open]);

  return (
    <div className="cellix-mode-dropdown-wrap" ref={wrapRef}>
      <button
        type="button"
        className={`cellix-mode-trigger ${open ? 'open' : ''}`}
        onClick={() => !disabled && setOpen((prev) => !prev)}
        disabled={disabled}
        title={activeMeta.hint}
        aria-haspopup="menu"
        aria-expanded={open}
      >
        <ActiveIcon size={13} />
        <span>{activeMeta.label}</span>
        <ChevronDown size={12} className="cellix-mode-chevron" />
      </button>

      {open && (
        <div className="cellix-mode-menu" role="menu" aria-label="Assistant mode">
          {ASSISTANT_MODES.map((m) => {
            const meta = ASSISTANT_MODE_META[m];
            const active = m === mode;
            const OptionIcon = modeIcon[m];

            return (
              <button
                key={m}
                type="button"
                role="menuitemradio"
                aria-checked={active}
                className={`cellix-mode-menu-item ${active ? 'active' : ''}`}
                onClick={() => {
                  onModeChange(m);
                  setOpen(false);
                }}
              >
                <OptionIcon size={13} />
                <span className="cellix-mode-menu-copy">
                  <span className="cellix-mode-menu-label">{meta.label}</span>
                  <span className="cellix-mode-menu-hint">{meta.hint}</span>
                </span>
                {active && <Check size={13} className="cellix-mode-check" />}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
};

const SUGGESTIONS = [
  "What's in cell A1?",
  'Calculate the total',
  'Explain this sheet to me',
];

const cellixWorkflowIllustration = new URL(
  '../../assets/cellix-workflow-illustration.png',
  import.meta.url,
).href;

interface EmptyStateProps {
  onSuggestion: (text: string) => void;
  children?: React.ReactNode;
}

export const EmptyState: React.FC<EmptyStateProps> = ({ onSuggestion, children }) => (
  <div className="cellix-empty">
    <img
      className="cellix-empty-illustration"
      src={cellixWorkflowIllustration}
      alt="Cellix workbook automation illustration"
    />
    <div className="cellix-empty-intro" aria-label="Cellix starter prompt">
      <div>
        <h2>
          Hi there, <span>User</span>
          <br />
          What would you like to <span>review?</span>
        </h2>
        <p>Review GST data, match ledgers, validate invoices, or prepare audit-ready Excel work.</p>
      </div>
    </div>
    {children}
    <div className="cellix-suggestion-carousel" aria-label="Suggested prompts">
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
  const [attachedFiles, setAttachedFiles] = React.useState<File[]>([]);
  const [composerSegments, setComposerSegments] = React.useState<ComposerSegment[]>([]);
  const [referenceOptions, setReferenceOptions] = React.useState<ReferenceOption[]>([]);
  const [referenceMenuOpen, setReferenceMenuOpen] = React.useState(false);
  const [referenceQuery, setReferenceQuery] = React.useState('');
  const [referenceStart, setReferenceStart] = React.useState<number | null>(null);
  const [activeReferenceIndex, setActiveReferenceIndex] = React.useState(0);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const referenceMenuRef = useRef<HTMLDivElement>(null);
  const nextSegmentIdRef = useRef(0);

  const createSegmentId = () => {
    nextSegmentIdRef.current += 1;
    return `segment-${nextSegmentIdRef.current}`;
  };

  const loadReferenceOptions = React.useCallback(async () => {
    if (typeof Excel === 'undefined') {
      setReferenceOptions([]);
      return;
    }

    try {
      await Excel.run(async (ctx) => {
        const worksheets = ctx.workbook.worksheets;
        worksheets.load('items/name,items/visibility');

        const activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
        activeWorksheet.load('name');

        const selectedRange = ctx.workbook.getSelectedRange();
        selectedRange.load('address');

        await ctx.sync();

        const visibleSheets = worksheets.items
          .filter((sheet) => sheet.visibility === Excel.SheetVisibility.visible)
          .map((sheet) => sheet.name);

        const options: ReferenceOption[] = [];
        options.push({
          id: 'group:workbook',
          label: 'Workbook',
          detail: `${visibleSheets.length} sheets`,
          insertText: '@workbook',
          type: 'group',
        });

        const selectedAddress = selectedRange.address;
        if (selectedAddress) {
          options.push({
            id: `range:${selectedAddress}`,
            label: 'Selected range',
            detail: selectedAddress,
            insertText: `@${selectedAddress}`,
            type: 'range',
          });
        }

        if (activeWorksheet.name) {
          options.push({
            id: `active:${activeWorksheet.name}`,
            label: 'Active sheet',
            detail: activeWorksheet.name,
            insertText: `@[${activeWorksheet.name}]`,
            type: 'sheet',
          });
        }

        visibleSheets.forEach((sheetName) => {
          options.push({
            id: `sheet:${sheetName}`,
            label: sheetName,
            detail: 'Sheet',
            insertText: `@[${sheetName}]`,
            type: 'sheet',
          });
        });

        setReferenceOptions(options);
      });
    } catch (error) {
      console.warn('[Cellix] Failed to load reference options:', error);
      setReferenceOptions([]);
    }
  }, []);

  const updateReferenceState = React.useCallback(
    (value: string, cursorPosition: number) => {
      const beforeCursor = value.slice(0, cursorPosition);
      const match = beforeCursor.match(/(^|\s)@([^\s@]*)$/);

      if (!match) {
        setReferenceMenuOpen(false);
        setReferenceQuery('');
        setReferenceStart(null);
        return;
      }

      const atIndex = beforeCursor.lastIndexOf('@');
      setReferenceStart(atIndex);
      setReferenceQuery(match[2] ?? '');
      setReferenceMenuOpen(true);
      setActiveReferenceIndex(0);
      void loadReferenceOptions();
    },
    [loadReferenceOptions],
  );

  const filteredReferenceOptions = React.useMemo(() => {
    const query = referenceQuery.trim().toLowerCase();
    if (!query) return referenceOptions;

    return referenceOptions.filter((option) => {
      return (
        option.label.toLowerCase().includes(query) ||
        option.detail.toLowerCase().includes(query) ||
        option.insertText.toLowerCase().includes(query)
      );
    });
  }, [referenceOptions, referenceQuery]);

  useEffect(() => {
    if (activeReferenceIndex >= filteredReferenceOptions.length) {
      setActiveReferenceIndex(0);
    }
  }, [activeReferenceIndex, filteredReferenceOptions.length]);

  useEffect(() => {
    if (!referenceMenuOpen) return;

    const handlePointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (
        referenceMenuRef.current?.contains(target) ||
        textareaRef.current?.contains(target)
      ) {
        return;
      }

      setReferenceMenuOpen(false);
    };

    document.addEventListener('mousedown', handlePointerDown);
    return () => document.removeEventListener('mousedown', handlePointerDown);
  }, [referenceMenuOpen]);

  const handleSend = () => {
    const segmentText = composerSegments
      .map((segment) => (segment.type === 'reference' ? segment.reference.insertText : segment.text))
      .join(' ');
    const composedMessage = [segmentText, message.trim()].filter(Boolean).join(' ');

    if (composedMessage && !disabled) {
      onSend(composedMessage);
      setMessage('');
      setAttachedFiles([]);
      setComposerSegments([]);
      if (textareaRef.current) textareaRef.current.style.height = 'auto';
    }
  };

  const handleFileSelect = (files: FileList | null) => {
    if (!files?.length) return;

    setAttachedFiles((current) => {
      const existingKeys = new Set(
        current.map((file) => `${file.name}:${file.size}:${file.lastModified}`),
      );
      const nextFiles = Array.from(files).filter((file) => {
        const key = `${file.name}:${file.size}:${file.lastModified}`;
        return !existingKeys.has(key);
      });

      return [...current, ...nextFiles];
    });

    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const removeAttachedFile = (fileToRemove: File) => {
    setAttachedFiles((current) => current.filter((file) => file !== fileToRemove));
  };

  const removeComposerSegment = (segmentId: string) => {
    setComposerSegments((current) => current.filter((segment) => segment.id !== segmentId));
  };

  const insertReference = (option: ReferenceOption) => {
    const textarea = textareaRef.current;
    const cursorPosition = textarea?.selectionStart ?? message.length;
    const start = referenceStart ?? cursorPosition;
    const beforeReference = message.slice(0, start).replace(/\s+$/, '');
    const afterReference = message.slice(cursorPosition).replace(/^\s+/, '');

    setMessage(afterReference);
    setComposerSegments((current) => {
      const next = [...current];
      if (beforeReference) {
        next.push({
          id: createSegmentId(),
          type: 'text',
          text: beforeReference,
        });
      }

      next.push({
        id: createSegmentId(),
        type: 'reference',
        reference: option,
      });

      return next;
    });
    setReferenceMenuOpen(false);
    setReferenceQuery('');
    setReferenceStart(null);

    window.setTimeout(() => {
      textareaRef.current?.focus();
      textareaRef.current?.setSelectionRange(0, 0);
    }, 0);
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Backspace' && !message && composerSegments.length > 0) {
      e.preventDefault();
      setComposerSegments((current) => {
        const last = current[current.length - 1];
        const next = current.slice(0, -1);

        if (last?.type === 'text') {
          setMessage(last.text);
          window.setTimeout(() => {
            const end = last.text.length;
            textareaRef.current?.setSelectionRange(end, end);
          }, 0);
        }

        return next;
      });
      return;
    }

    if (referenceMenuOpen) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setActiveReferenceIndex((index) =>
          filteredReferenceOptions.length === 0 ? 0 : (index + 1) % filteredReferenceOptions.length,
        );
        return;
      }

      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setActiveReferenceIndex((index) =>
          filteredReferenceOptions.length === 0
            ? 0
            : (index - 1 + filteredReferenceOptions.length) % filteredReferenceOptions.length,
        );
        return;
      }

      if (e.key === 'Escape') {
        e.preventDefault();
        setReferenceMenuOpen(false);
        return;
      }

      if (e.key === 'Enter' && filteredReferenceOptions[activeReferenceIndex]) {
        e.preventDefault();
        insertReference(filteredReferenceOptions[activeReferenceIndex]);
        return;
      }
    }

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
  const visibleMessageLength = Math.max(message.trimEnd().length + 1, 4);
  const textareaInlineSize = message
    ? composerSegments.length === 0
      ? '100%'
      : `${Math.min(visibleMessageLength, 80)}ch`
    : undefined;

  return (
    <div className="cellix-input-area">
      <div
        className={`cellix-input-shell mode-${mode} ${isProcessing ? 'processing' : ''} ${
          isWaitingClarification ? 'clarifying' : ''
        }`}
      >
        <div className="cellix-reference-editor" onClick={() => textareaRef.current?.focus()}>
          {composerSegments.map((segment) => {
            if (segment.type === 'text') {
              return (
                <span key={segment.id} className="cellix-composer-text-segment">
                  {segment.text}
                </span>
              );
            }

            const reference = segment.reference;
            const Icon =
              reference.type === 'range' ? AtSign : reference.type === 'group' ? Folder : Grid3X3;

            return (
              <div
                key={segment.id}
                className="cellix-reference-chip"
                title={`${reference.label} ${reference.detail}`}
              >
                <Icon size={12} />
                <span>{reference.label}</span>
                <button
                  type="button"
                  onClick={(event) => {
                    event.stopPropagation();
                    removeComposerSegment(segment.id);
                  }}
                  aria-label={`Remove ${reference.label}`}
                  disabled={inputDisabled}
                >
                  <X size={11} />
                </button>
              </div>
            );
          })}
          <textarea
            ref={textareaRef}
            className={message ? 'has-text' : ''}
            style={{
              flexBasis: textareaInlineSize,
              width: textareaInlineSize,
            }}
            value={message}
            onChange={(e) => {
              setMessage(e.target.value);
              e.target.style.height = 'auto';
              e.target.style.height = `${Math.min(e.target.scrollHeight, 120)}px`;
              updateReferenceState(e.target.value, e.target.selectionStart);
            }}
            onClick={(e) => updateReferenceState(message, e.currentTarget.selectionStart)}
            onSelect={(e) => updateReferenceState(message, e.currentTarget.selectionStart)}
            onKeyDown={handleKeyDown}
            placeholder={composerSegments.length > 0 ? '' : resolvedPlaceholder}
            disabled={inputDisabled}
            rows={1}
          />
        </div>
        {referenceMenuOpen && (
          <div
            className="cellix-reference-menu"
            ref={referenceMenuRef}
            role="listbox"
            aria-label="Workbook references"
          >
            {filteredReferenceOptions.length > 0 ? (
              filteredReferenceOptions.map((option, index) => {
                const Icon =
                  option.type === 'range' ? AtSign : option.type === 'group' ? Folder : Grid3X3;

                return (
                  <button
                    key={option.id}
                    type="button"
                    role="option"
                    aria-selected={index === activeReferenceIndex}
                    className={`cellix-reference-option ${
                      index === activeReferenceIndex ? 'active' : ''
                    }`}
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => insertReference(option)}
                  >
                    <Icon size={13} />
                    <span className="cellix-reference-copy">
                      <span className="cellix-reference-label">{option.label}</span>
                      <span className="cellix-reference-detail">{option.detail}</span>
                    </span>
                  </button>
                );
              })
            ) : (
              <div className="cellix-reference-empty">No references found</div>
            )}
          </div>
        )}
        {attachedFiles.length > 0 && (
          <div className="cellix-attachment-list" aria-label="Attached documents">
            {attachedFiles.map((file) => (
              <div
                key={`${file.name}:${file.size}:${file.lastModified}`}
                className="cellix-attachment-chip"
                title={file.name}
              >
                <FileText size={13} />
                <span>{file.name}</span>
                <button
                  type="button"
                  onClick={() => removeAttachedFile(file)}
                  aria-label={`Remove ${file.name}`}
                  disabled={inputDisabled}
                >
                  <X size={12} />
                </button>
              </div>
            ))}
          </div>
        )}
        <div className="cellix-input-toolbar">
          <div className="cellix-input-toolbar-left">
            <ModeSwitch mode={mode} onModeChange={onModeChange} disabled={isProcessing} />
            <button
              type="button"
              className="cellix-upload-doc-btn"
              onClick={() => fileInputRef.current?.click()}
              disabled={inputDisabled || isProcessing}
              aria-label="Upload documents"
              title="Upload documents"
            >
              <Paperclip size={14} />
            </button>
            <input
              ref={fileInputRef}
              type="file"
              className="cellix-upload-doc-input"
              accept=".csv,.doc,.docx,.pdf,.txt,.xls,.xlsx"
              multiple
              onChange={(event) => handleFileSelect(event.target.files)}
            />
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
              disabled={(!message.trim() && composerSegments.length === 0) || inputDisabled}
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
  sessions: import('@/types/chatSession').ChatSession[];
  activeSessionId: string | null;
  turns: import('@/types/conversationTurn').ConversationTurn[];
  activeTurnId: string | null;
  isWaitingForResponse: boolean;
  isWaitingClarification: boolean;
  activeClarification: ClarificationPayload | null;
  previewEnabled: boolean;
  mode: AssistantMode;
  onModeChange: (mode: AssistantMode) => void;
  onRunAsAction: (message: string) => void;
  compareResult: CompareResult | null;
  isComparing: boolean;
  onCloseCompare: () => void;
  onSend: (message: string) => void;
  onStop: () => void;
  onNewChat: () => void;
  onSelectSession: (sessionId: string) => void;
  onCloseSession: (sessionId: string) => void;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onAnswerQuestion: (answer: string) => void;
  onClarificationAnswer: (answer: string) => void;
  onClarificationDismiss: () => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onFollowUp: (text: string) => void;
  conversationId: string | null;
  onRevertHistoryEntry: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
  isApplyingActions?: boolean;
  pendingPreview?: {
    changes: CellChange[];
    changeSetId?: string;
    summary: string;
    isApplying: boolean;
    onAccept: () => void;
    onReject: () => void;
  } | null;
  refinementChangeSetId?: string | null;
  quickEditMode?: boolean;
  onStartQuickEdit?: () => void;
  onCancelQuickEdit?: () => void;
}

const ConversationPanel: React.FC<ConversationPanelProps> = ({
  sessions,
  activeSessionId,
  turns,
  activeTurnId,
  isWaitingForResponse,
  isWaitingClarification,
  activeClarification,
  previewEnabled,
  mode,
  onModeChange,
  onRunAsAction,
  compareResult,
  isComparing,
  onCloseCompare,
  onSend,
  onStop,
  onNewChat,
  onSelectSession,
  onCloseSession,
  onAcceptActions,
  onRejectActions,
  onAnswerQuestion,
  onClarificationAnswer,
  onClarificationDismiss,
  onToggleThinking,
  onAnswerComplete,
  onFollowUp,
  conversationId,
  onRevertHistoryEntry,
  isApplyingActions = false,
  pendingPreview = null,
  refinementChangeSetId = null,
  quickEditMode = false,
  onStartQuickEdit,
  onCancelQuickEdit,
}) => {
  const contentRef = useRef<HTMLDivElement>(null);
  const showStartScreen = turns.length === 0;
  const activeTurn = activeTurnId ? turns.find((turn) => turn.id === activeTurnId) : undefined;
  const previewActionsReady = Boolean(
    activeTurn && !isWaitingForResponse && isTurnPresentationComplete(activeTurn),
  );
  const handleQuestionAnswer = (answer: string) => {
    if (activeClarification) {
      onClarificationAnswer(answer);
      return;
    }
    onAnswerQuestion(answer);
  };
  const previewItems: DiffItem[] = pendingPreview
    ? pendingPreview.changes.map((change) => ({
        sheetName: change.sheet,
        address: change.cell,
        actionType: change.formula ? 'SET_FORMULA' : 'SET_CELL',
        before: formatCellValue(change.before),
        after: formatCellValue(change.after),
        description: change.formula
          ? `Set formula ${change.formula}`
          : `Change value to ${formatCellValue(change.after)}`,
      }))
    : [];

  useEffect(() => {
    contentRef.current?.scrollTo({ top: contentRef.current.scrollHeight, behavior: 'smooth' });
  }, [turns, activeTurnId, isWaitingForResponse, activeClarification]);

  return (
    <div className="cellix-panel">
      <PanelHeader
        sessions={sessions}
        activeSessionId={activeSessionId}
        isWaitingForResponse={isWaitingForResponse}
        onSelectSession={onSelectSession}
        onCloseSession={onCloseSession}
        onNewChat={onNewChat}
      />

      <div className={`cellix-content ${showStartScreen ? 'start' : ''}`} ref={contentRef}>
        {showStartScreen ? (
          <EmptyState onSuggestion={onSend}>
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
                    ? 'Ask anything about your workbook - use @ for references'
                    : mode === 'plan'
                      ? 'Describe what you want to plan - use @ for references'
                      : 'Describe the change you want to make - use @ for references'
              }
            />
          </EmptyState>
        ) : (
          turns.map((turn) => (
            <TurnRenderer
              key={turn.id}
              turn={turn}
              isActive={turn.id === activeTurnId}
              isWaiting={isWaitingForResponse && turn.id === activeTurnId}
              previewEnabled={previewEnabled}
              isApplying={isApplyingActions}
              showActionButtons={previewActionsReady}
              onAcceptActions={onAcceptActions}
              onRejectActions={onRejectActions}
              onAnswerQuestion={handleQuestionAnswer}
              onToggleThinking={onToggleThinking}
              onAnswerComplete={onAnswerComplete}
              onFollowUp={onFollowUp}
              onRunAsAction={onRunAsAction}
            />
          ))
        )}
      </div>

      {(isComparing || compareResult) && (
        <div className="cellix-bottom-tools">
          {(isComparing || compareResult) && (
            <SheetCompareView
              result={compareResult}
              isLoading={isComparing}
              onClose={onCloseCompare}
            />
          )}
        </div>
      )}

      {pendingPreview && (
        <PreviewSummaryBar
          items={previewItems}
          summary={pendingPreview.summary}
          onAccept={pendingPreview.onAccept}
          onReject={pendingPreview.onReject}
          isApplying={pendingPreview.isApplying}
          showActions={previewActionsReady}
        />
      )}

      <ChangeHistoryPanel conversationId={conversationId} onRevert={onRevertHistoryEntry} />

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

      {!showStartScreen && (
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
                ? 'Ask anything about your workbook - use @ for references'
                : mode === 'plan'
                  ? 'Describe what you want to plan - use @ for references'
                  : 'Describe the change you want to make - use @ for references'
          }
        />
      )}
    </div>
  );
};

export default ConversationPanel;
