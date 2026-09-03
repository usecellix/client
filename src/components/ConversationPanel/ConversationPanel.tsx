import React, { useRef, useEffect, KeyboardEvent } from 'react';
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
import { isTurnPresentationComplete } from '@/utils/turnPresentation';
import { LastChangeRevert } from '@/components/ChangeHistoryPanel/LastChangeRevert';
import { RestoreResult } from '@/types/checkpoint';
import { SheetAction } from '@/types/sheet-actions';
import { useSession } from '@/auth/auth-client';
import { TextAnimate } from '@/components/ui/text-animate';
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

const HEADING_CHAR_DURATION = 0.22;

const getDayGreeting = () => {
  // Always use IST (Asia/Kolkata), not the browser's local timezone.
  const hour = Number(
    new Intl.DateTimeFormat('en-GB', {
      timeZone: 'Asia/Kolkata',
      hour: 'numeric',
      hourCycle: 'h23',
    }).format(new Date()),
  );
  if (hour < 12) return 'Good morning';
  if (hour < 16) return 'Good afternoon';
  return 'Good evening';
};

/** When the next character-stagger TextAnimate should start for a continuous feel. */
const getNextCharDelay = (delay = 0) => delay + HEADING_CHAR_DURATION;

export const EmptyState: React.FC<EmptyStateProps> = ({ onSuggestion, children }) => {
  const { data: session } = useSession();
  const userName = session?.user?.name?.trim().split(/\s+/)[0] || 'User';
  const greetingPrefix = React.useMemo(() => `${getDayGreeting()}, `, []);
  const promptPrefix = 'What would you like to ';
  const promptAccent = 'review?';
  const nameDelay = getNextCharDelay();
  const promptDelay = getNextCharDelay(nameDelay) + 0.08;
  const accentDelay = getNextCharDelay(promptDelay);

  return (
    <div className="cellix-empty">
      <img
        className="cellix-empty-illustration"
        src={cellixWorkflowIllustration}
        alt="Cellix workbook automation illustration"
      />
      <div className="cellix-empty-intro" aria-label="Cellix starter prompt">
        <div>
          <h2>
            <span className="cellix-empty-heading-line">
              <TextAnimate
                as="span"
                animation="blurInUp"
                by="character"
                once
                startOnView={false}
                duration={HEADING_CHAR_DURATION}
                className="cellix-empty-heading-animate"
              >
                {greetingPrefix}
              </TextAnimate>
              <TextAnimate
                as="span"
                animation="blurInUp"
                by="character"
                once
                startOnView={false}
                delay={nameDelay}
                duration={HEADING_CHAR_DURATION}
                className="cellix-empty-heading-animate cellix-empty-accent"
              >
                {userName}
              </TextAnimate>
            </span>
            <span className="cellix-empty-heading-line">
              <TextAnimate
                as="span"
                animation="blurInUp"
                by="character"
                once
                startOnView={false}
                delay={promptDelay}
                duration={HEADING_CHAR_DURATION}
                className="cellix-empty-heading-animate"
              >
                {promptPrefix}
              </TextAnimate>
              <TextAnimate
                as="span"
                animation="blurInUp"
                by="character"
                once
                startOnView={false}
                delay={accentDelay}
                duration={HEADING_CHAR_DURATION}
                className="cellix-empty-heading-animate cellix-empty-accent"
              >
                {promptAccent}
              </TextAnimate>
            </span>
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
};

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
  /** Rename an open tab, and its server conversation if it has one (TASKS.md #177). */
  onRenameSession: (sessionId: string, title: string) => void;
  /** Delete an open tab, and its server conversation if it has one (TASKS.md #177). */
  onDeleteSession: (sessionId: string) => Promise<void>;
  /** Delete a history row that isn't necessarily an open tab (TASKS.md #177). */
  onDeleteHistoryConversation: (conversationId: string) => Promise<void>;
  /** Open a past conversation from server-backed history (TASKS.md #172). */
  onOpenHistoryConversation: (conversationId: string) => Promise<boolean>;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onAcceptAllActions?: (turnId: string, fromBlockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onAnswerQuestion: (answer: string) => void;
  onClarificationAnswer: (answer: string) => void;
  onClarificationDismiss: () => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onFollowUp: (text: string) => void;
  onRegenerate?: (turnId: string, overrideMessage?: string) => void;
  conversationId: string | null;
  onRevertHistoryEntry: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
  workbookId?: string;
  onRestoreCheckpoint: (result: RestoreResult) => Promise<void>;
  isApplyingActions?: boolean;
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
  onRenameSession,
  onDeleteSession,
  onDeleteHistoryConversation,
  onOpenHistoryConversation,
  onAcceptActions,
  onAcceptAllActions,
  onRejectActions,
  onAnswerQuestion,
  onClarificationAnswer,
  // onClarificationDismiss is supplied by App but no control invokes it yet —
  // the dismiss affordance is unimplemented, so it is intentionally not bound.
  onToggleThinking,
  onAnswerComplete,
  onFollowUp,
  onRegenerate,
  conversationId,
  onRevertHistoryEntry,
  workbookId,
  onRestoreCheckpoint,
  isApplyingActions = false,
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

  // Sticky-to-bottom, like Claude/Cursor/Codex chat panes: only auto-scroll
  // when the user is already at (or near) the bottom, or when a brand-new
  // turn just landed. Without this, any in-place edit to existing content —
  // expanding "Thought process" on an old turn, answering a clarification,
  // Accept/Reject — replaces `turns` with a new array reference and yanked
  // the whole view down to the latest message, even far above the fold.
  const isNearBottomRef = useRef(true);
  const prevTurnsLengthRef = useRef(turns.length);

  const handleContentScroll = () => {
    const el = contentRef.current;
    if (!el) return;
    const distanceFromBottom = el.scrollHeight - el.scrollTop - el.clientHeight;
    isNearBottomRef.current = distanceFromBottom < 80;
  };

  useEffect(() => {
    const el = contentRef.current;
    if (!el) return;
    const turnAppended = turns.length > prevTurnsLengthRef.current;
    prevTurnsLengthRef.current = turns.length;
    // A newly appended turn (the user just sent something) always jumps to
    // it — matching the intentional "show me what I just did" case. Any
    // other change (streaming reveal, in-place toggles) only follows along
    // if the user hadn't already scrolled away to read something else.
    if (turnAppended || isNearBottomRef.current) {
      el.scrollTo({ top: el.scrollHeight, behavior: 'smooth' });
      isNearBottomRef.current = true;
    }
  }, [turns, activeTurnId, isWaitingForResponse, activeClarification]);

  const composerInput = (
    <PanelInput
      onSend={onSend}
      onStop={onStop}
      disabled={isWaitingForResponse}
      isProcessing={isWaitingForResponse}
      isWaitingClarification={isWaitingClarification}
      mode={mode}
      onModeChange={onModeChange}
      placeholder={
        mode === 'ask'
          ? 'Ask anything about your workbook - use @ for references'
          : mode === 'plan'
            ? 'Describe what you want to plan - use @ for references'
            : 'Describe the change you want to make - use @ for references'
      }
    />
  );

  const composerDock = (
    <div className="cellix-composer-dock">
      {composerInput}
    </div>
  );

  return (
    <div className="cellix-panel">
      <PanelHeader
        sessions={sessions}
        activeSessionId={activeSessionId}
        isWaitingForResponse={isWaitingForResponse}
        onSelectSession={onSelectSession}
        onCloseSession={onCloseSession}
        onRenameSession={onRenameSession}
        onDeleteSession={onDeleteSession}
        onDeleteHistoryConversation={onDeleteHistoryConversation}
        onNewChat={onNewChat}
        onOpenHistoryConversation={onOpenHistoryConversation}
        // Checkpoints icon temporarily hidden (not removed) — feature, panel,
        // and backend are all still intact behind this flag. Restore with
        // `showCheckpointsButton={!showStartScreen}` when it's wanted again.
        showCheckpointsButton={false}
        workbookId={workbookId}
        conversationId={conversationId}
        onRestoreCheckpoint={onRestoreCheckpoint}
      />

      <div
        className={`cellix-content ${showStartScreen ? 'start' : ''}`}
        ref={contentRef}
        onScroll={handleContentScroll}
      >
        {showStartScreen ? (
          <EmptyState onSuggestion={onSend}>{composerDock}</EmptyState>
        ) : (
          turns.map((turn) => (
            <TurnRenderer
              key={turn.id}
              turn={turn}
              isActive={turn.id === activeTurnId}
              isWaiting={isWaitingForResponse && turn.id === activeTurnId}
              previewEnabled={previewEnabled}
              isApplying={isApplyingActions}
              showActionButtons={
                turn.id === activeTurnId
                  ? previewActionsReady
                  : isTurnPresentationComplete(turn)
              }
              onAcceptActions={onAcceptActions}
              onAcceptAllActions={onAcceptAllActions}
              onRejectActions={onRejectActions}
              onAnswerQuestion={handleQuestionAnswer}
              onToggleThinking={onToggleThinking}
              onAnswerComplete={onAnswerComplete}
              onFollowUp={onFollowUp}
              onRegenerate={onRegenerate}
              onRunAsAction={onRunAsAction}
              onRevertChangeSet={onRevertHistoryEntry}
            />
          ))
        )}

        {/* Revert for the most recent applied change lives at the end of the
            conversation, where the change just happened — not behind the
            composer's history icon. Full per-entry history is still there. */}
        {!showStartScreen && (
          <LastChangeRevert
            conversationId={conversationId}
            onRevert={onRevertHistoryEntry}
            refreshKey={turns.length}
          />
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

      {!showStartScreen && composerDock}
    </div>
  );
};

export default ConversationPanel;
