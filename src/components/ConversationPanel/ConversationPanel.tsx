import React, { useRef, useEffect, useMemo, useState, KeyboardEvent } from 'react';
import {
  AlertCircle,
  ArrowRight,
  AtSign,
  Check,
  ChevronDown,
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
import { RestoreResult } from '@/types/checkpoint';
import { SheetAction } from '@/types/sheet-actions';
import { TurnBlock } from '@/types/conversationTurn';
import { RawTable } from '@/types/rawTable';
import { ATTACHMENT_ACCEPT, attachmentExtension, draggedFileKind, DraggedFileKind } from '@/services/attachments/attachmentError';
import { useSession } from '@/auth/auth-client';
import { TextAnimate } from '@/components/ui/text-animate';
import PanelHeader from './PanelHeader';
import AttachmentChip from './AttachmentChip';
import FileTypeBadge from './FileTypeBadge';
import TurnRenderer from './TurnRenderer';
import QuestionChoicesPanel from './QuestionChoicesPanel';
import CreditUpgradeCard, { CREDIT_NOTICE_THRESHOLD } from './CreditUpgradeCard';
import ResumeRunCard from './ResumeRunCard';

/* global Excel */

export type ImportAttachmentHandler = (
  rawTable: RawTable,
  fileName: string,
  question?: string,
) => void | boolean | Promise<void | boolean>;

interface PanelInputProps {
  onSend: (message: string) => void;
  onStop?: () => void;
  disabled?: boolean;
  isProcessing?: boolean;
  isWaitingClarification?: boolean;
  placeholder?: string;
  mode: AssistantMode;
  onModeChange: (mode: AssistantMode) => void;
  /**
   * Imports an attached file that has been decoded in the task pane, with the
   * question the user typed alongside it, if any. The returned promise should
   * settle when that import's turn has finished, so the next attached file is
   * not started on top of it; resolving `false` means the import failed and
   * the question was not used. Without this prop, attached files are shown but
   * not imported.
   */
  onImportAttachment?: ImportAttachmentHandler;
  /**
   * Files dropped on the task pane. The panel owns the drop target (the whole
   * pane, not just the box), so it hands each drop over here, where it goes
   * through the same checks as the paperclip. `id` changes on every drop.
   */
  droppedFiles?: { id: number; files: File[] } | null;
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
  'Reconcile my GST purchases with GSTR-2B',
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
      <p className="cellix-empty-narrow-hint" role="note">
        <span aria-hidden="true">←</span> Expand panel to continue
      </p>
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
  onImportAttachment,
  droppedFiles,
}) => {
  const [message, setMessage] = React.useState('');
  const [attachedFiles, setAttachedFiles] = React.useState<File[]>([]);
  const [attachmentNotice, setAttachmentNotice] = React.useState<string | null>(null);
  // Files that have been read and are waiting for Send, and the counter that
  // tells their chips Send was pressed.
  const [readyAttachmentKeys, setReadyAttachmentKeys] = React.useState<string[]>([]);
  const [importRequest, setImportRequest] = React.useState(0);
  // What the user typed alongside the attached file. Handed to the first import
  // that Send starts, which asks it once the statement is in the workbook.
  const attachmentQuestionRef = useRef('');
  const [composerSegments, setComposerSegments] = React.useState<ComposerSegment[]>([]);
  const [referenceOptions, setReferenceOptions] = React.useState<ReferenceOption[]>([]);
  const [referenceMenuOpen, setReferenceMenuOpen] = React.useState(false);
  const [referenceQuery, setReferenceQuery] = React.useState('');
  const [referenceStart, setReferenceStart] = React.useState<number | null>(null);
  const [activeReferenceIndex, setActiveReferenceIndex] = React.useState(0);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  // True from the moment one attachment's import is handed off until its turn
  // ends. Files import one at a time; a second file attached meanwhile waits.
  const importInFlightRef = useRef(false);
  const referenceMenuRef = useRef<HTMLDivElement>(null);
  const nextSegmentIdRef = useRef(0);

  const inputDisabled = disabled || isWaitingClarification;

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

    // A file is attached: this Send imports it, and whatever was typed goes with
    // it as the question to answer once the statement is in the workbook.
    if (readyAttachmentKeys.length > 0 && !disabled) {
      attachmentQuestionRef.current = composedMessage;
      setImportRequest((count) => count + 1);
      setMessage('');
      setComposerSegments([]);
      if (textareaRef.current) textareaRef.current.style.height = 'auto';
      return;
    }

    if (composedMessage && !disabled) {
      onSend(composedMessage);
      setMessage('');
      setComposerSegments([]);
      if (textareaRef.current) textareaRef.current.style.height = 'auto';
    }
  };

  const handleFileSelect = (files: FileList | File[] | null) => {
    if (!files?.length) return;

    // The picker's `accept` list is only a hint the user can override, so the
    // type is checked again here rather than leaving a chip that cannot be read.
    const picked = Array.from(files);
    const supported = picked.filter((file) => attachmentExtension(file.name) !== null);
    const refused = picked.filter((file) => attachmentExtension(file.name) === null);
    setAttachmentNotice(
      refused.length
        ? `${refused.map((file) => file.name).join(', ')} cannot be attached. Use a PDF, XLSX, XLS or CSV file.`
        : null,
    );

    setAttachedFiles((current) => {
      const existingKeys = new Set(
        current.map((file) => `${file.name}:${file.size}:${file.lastModified}`),
      );
      const nextFiles = supported.filter((file) => {
        const key = `${file.name}:${file.size}:${file.lastModified}`;
        return !existingKeys.has(key);
      });

      return [...current, ...nextFiles];
    });

    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  // A drop on the pane is treated exactly like picking the file with the paperclip,
  // and is ignored whenever the paperclip is disabled.
  const handledDropRef = useRef<number | null>(null);
  useEffect(() => {
    if (!droppedFiles || handledDropRef.current === droppedFiles.id) return;
    handledDropRef.current = droppedFiles.id;
    if (inputDisabled || isProcessing) return;
    handleFileSelect(droppedFiles.files);
    // handleFileSelect is recreated every render; a drop is handled once, by id.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [droppedFiles]);

  const removeAttachedFile = (fileToRemove: File) => {
    setAttachedFiles((current) => current.filter((file) => file !== fileToRemove));
  };

  const attachmentKey = (file: File) => `${file.name}:${file.size}:${file.lastModified}`;

  const setAttachmentReady = (file: File, ready: boolean) => {
    const key = attachmentKey(file);
    setReadyAttachmentKeys((current) => {
      const has = current.includes(key);
      if (ready === has) return current;
      return ready ? [...current, key] : current.filter((k) => k !== key);
    });
    // The file is read (and unlocked): the next thing the user does is type a
    // question or press Enter, both of which happen in the message box.
    if (ready) textareaRef.current?.focus();
  };

  /** Starts one file's import. False means "not now": the chip keeps the file and offers it again. */
  const startAttachmentImport = (file: File, rawTable: RawTable): boolean => {
    if (!onImportAttachment || importInFlightRef.current || inputDisabled || isProcessing) return false;
    importInFlightRef.current = true;
    removeAttachedFile(file);
    const question = attachmentQuestionRef.current;
    attachmentQuestionRef.current = '';
    void Promise.resolve()
      .then(() => onImportAttachment(rawTable, file.name, question || undefined))
      .then((ok) => {
        // The import failed, so the question was never asked. Give it back
        // rather than lose what the user typed, unless they have typed since.
        if (ok === false && question) setMessage((current) => current || question);
      })
      .catch((error) => console.error('[Cellix] Attachment import failed to start:', error))
      .finally(() => {
        importInFlightRef.current = false;
      });
    return true;
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

  const resolvedPlaceholder = isWaitingClarification
    ? '⏸ Answer the question above first…'
    : isProcessing
      ? 'Processing your request…'
      : readyAttachmentKeys.length > 0
        ? 'Ask about this statement, or press Enter to import it'
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
        }${attachedFiles.length > 0 ? ' has-attachments' : ''}`}
      >
        {attachedFiles.length > 0 && (
          <div className="cellix-attachment-list" aria-label="Attached documents">
            {attachedFiles.map((file) => (
              <AttachmentChip
                key={attachmentKey(file)}
                file={file}
                disabled={inputDisabled || isProcessing}
                onRemove={() => removeAttachedFile(file)}
                onImport={
                  onImportAttachment ? (rawTable) => startAttachmentImport(file, rawTable) : undefined
                }
                importRequest={importRequest}
                onReadyChange={(ready) => setAttachmentReady(file, ready)}
              />
            ))}
          </div>
        )}
        {attachmentNotice && (
          <div className="cellix-attachment-note is-error cellix-attachment-notice" role="alert">
            {attachmentNotice}
          </div>
        )}
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
              accept={ATTACHMENT_ACCEPT}
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
              disabled={
                (!message.trim() && composerSegments.length === 0 && readyAttachmentKeys.length === 0) ||
                inputDisabled
              }
              aria-label={readyAttachmentKeys.length > 0 ? 'Send with the attached statement' : 'Send'}
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
  /** Imports an attached file once the task pane has decoded it, with any question typed alongside. */
  onImportAttachment?: ImportAttachmentHandler;
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
  /** Finish a step that failed for want of a sheet. TASKS.md #346. */
  onRetryMissingSheets?: (turnId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onResolveGstReconCollision: (
    turnId: string,
    blockId: string,
    collisionId: string,
    choice: 'overwrite' | 'new',
  ) => void;
  onRejectGstReconCollision: (turnId: string, blockId: string, collisionId: string) => void;
  onJumpToGstReconRow: (sheetName: string, row: number) => void;
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
  creditAccount?: import('@/services/billingService').CreditAccountSummary | null;
  isLowCreditBalance?: boolean;
  /**
   * An unfinished build this conversation can carry on from, or null.
   * LONG_PROMPT_RELIABILITY_PLAN.md Phase 8.
   */
  resumableRun?: { runId: string; waveIndex: number; waveTotal: number } | null;
  onResumeRun?: () => void;
  onDismissResumableRun?: () => void;
}

const ConversationPanel: React.FC<ConversationPanelProps> = ({
  onImportAttachment,
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
  onRetryMissingSheets,
  onRejectActions,
  onResolveGstReconCollision,
  onRejectGstReconCollision,
  onJumpToGstReconRow,
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
  creditAccount = null,
  isLowCreditBalance = false,
  resumableRun = null,
  onResumeRun,
  onDismissResumableRun,
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

  // The one question actually awaiting an answer, lifted out of the transcript
  // so it can be docked directly above the composer the way Cursor and Claude
  // Code ask theirs. A question buried mid-scroll reads as history: the user
  // has to find it before they can answer it, and the composer underneath sits
  // disabled with no visible reason. Docking it puts the question, its choices,
  // and the input the user would otherwise type into in one place.
  //
  // Only ever one: the newest turn still in `awaiting_input`. Older turns keep
  // their answered questions inline as part of the record. TASKS.md #182.
  // Dismissed for this session only — the balance keeps falling, so a nudge
  // permanently silenced on one click would go quiet exactly when it starts to
  // matter. TASKS.md #199.
  const [creditNoticeDismissed, setCreditNoticeDismissed] = useState(false);

  // Lifted out of QuestionChoicesPanel so the optional-details composer can
  // submit the chosen option and the typed detail as one answer. TASKS.md #186.
  const [questionSelection, setQuestionSelection] = useState<string | null>(null);

  const pendingQuestion = useMemo(() => {
    for (let i = turns.length - 1; i >= 0; i -= 1) {
      const turn = turns[i];
      if (turn.phase !== 'awaiting_input') continue;
      const block = turn.blocks.find(
        (b): b is Extract<TurnBlock, { type: 'question' }> =>
          b.type === 'question' && b.revealState !== 'hidden' && !b.answeredWith,
      );
      if (block) return block;
    }
    return null;
  }, [turns]);

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

  // A new question must never inherit the previous one's selection.
  useEffect(() => {
    setQuestionSelection(null);
  }, [pendingQuestion?.id]);

  // Dismissing is a "not now", not a "never". Bring the card back the moment
  // the balance actually crosses into low — that is a different, newly urgent
  // fact from the healthy balance the user waved away, and it is the one
  // moment the nudge exists for. TASKS.md #203.
  const creditBelowNoticeThreshold =
    creditAccount != null && creditAccount.availableBalance < CREDIT_NOTICE_THRESHOLD;

  // Dismissing hides the card for now; dropping under the threshold again brings it back.
  useEffect(() => {
    if (isLowCreditBalance || creditBelowNoticeThreshold) setCreditNoticeDismissed(false);
  }, [isLowCreditBalance, creditBelowNoticeThreshold]);

  // Drag a file onto the pane to attach it. `dragDepth` counts enter and leave
  // events, because moving over a child element fires a leave on its parent.
  const [dropFiles, setDropFiles] = useState<{ id: number; files: File[] } | null>(null);
  const [draggedKinds, setDraggedKinds] = useState<DraggedFileKind[] | null>(null);
  const draggedKindsRef = useRef<DraggedFileKind[] | null>(null);
  const isDraggingFile = draggedKinds !== null;
  const dragDepthRef = useRef(0);
  const dragCarriesFiles = (event: React.DragEvent) => Array.from(event.dataTransfer?.types ?? []).includes('Files');
  const draggedKindsOf = (event: React.DragEvent): DraggedFileKind[] =>
    Array.from(event.dataTransfer?.items ?? [])
      .filter((item) => item.kind === 'file')
      .map((item) => draggedFileKind(item.type));
  const showDraggedKinds = (kinds: DraggedFileKind[] | null) => {
    draggedKindsRef.current = kinds;
    setDraggedKinds(kinds);
  };
  // Only a drag made up entirely of files that can never be attached is refused
  // at the cursor; a mixed or uncertain drag is let through and sorted on drop.
  const dragIsAllUnsupported = (kinds: DraggedFileKind[] | null) =>
    !!kinds && kinds.length > 0 && kinds.every((kind) => kind === 'unsupported');

  const handlePaneDragEnter = (event: React.DragEvent) => {
    if (!dragCarriesFiles(event)) return;
    event.preventDefault();
    dragDepthRef.current += 1;
    if (!draggedKindsRef.current) showDraggedKinds(draggedKindsOf(event));
  };
  const handlePaneDragOver = (event: React.DragEvent) => {
    if (!dragCarriesFiles(event)) return;
    // Without this the webview opens the dropped file instead of handing it over.
    event.preventDefault();
    event.dataTransfer.dropEffect = dragIsAllUnsupported(draggedKindsRef.current) ? 'none' : 'copy';
  };
  const handlePaneDragLeave = (event: React.DragEvent) => {
    if (!dragCarriesFiles(event)) return;
    dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
    if (dragDepthRef.current === 0) showDraggedKinds(null);
  };
  const handlePaneDrop = (event: React.DragEvent) => {
    if (!dragCarriesFiles(event)) return;
    event.preventDefault();
    dragDepthRef.current = 0;
    showDraggedKinds(null);
    const files = Array.from(event.dataTransfer.files);
    if (files.length) setDropFiles((current) => ({ id: (current?.id ?? 0) + 1, files }));
  };

  const dragHasPdf = !!draggedKinds?.includes('pdf');
  const dragHasSheet = !!draggedKinds?.includes('sheet');
  const dragRefused = dragIsAllUnsupported(draggedKinds);
  const dropTitle = dragRefused
    ? "This file type can't be attached"
    : dragHasPdf && !dragHasSheet && draggedKinds?.length === 1
      ? 'Drop the PDF to attach it'
      : 'Drop to attach';

  const composerInput = (
    <PanelInput
      onSend={
        pendingQuestion
          ? (text) =>
              handleQuestionAnswer(
                questionSelection ? `${questionSelection} — ${text}` : text,
              )
          : onSend
      }
      onStop={onStop}
      onImportAttachment={onImportAttachment}
      droppedFiles={dropFiles}
      disabled={isWaitingForResponse}
      isProcessing={isWaitingForResponse}
      // The docked card is the thing asking; this field is optional detail
      // alongside it, so it stays live instead of being blocked by the flag.
      isWaitingClarification={pendingQuestion ? false : isWaitingClarification}
      mode={mode}
      onModeChange={onModeChange}
      placeholder={
        pendingQuestion
          ? 'Add more optional details'
          : mode === 'ask'
            ? 'Ask anything about your workbook - use @ for references'
            : mode === 'plan'
              ? 'Describe what you want to plan - use @ for references'
              : 'Describe the change you want to make - use @ for references'
      }
    />
  );

  // Question card above, composer below — Cursor's arrangement. #185 removed
  // the composer entirely because it could only render disabled behind a
  // "⏸ Answer the question above first…" nag; it comes back here as Cursor's
  // "Add more optional details" field, live rather than blocked, so a choice
  // can be qualified in prose instead of forcing Other. Sending from it
  // answers the question (folding in the selected option) rather than starting
  // a new turn. TASKS.md #186.
  // Shown only once the balance is under CREDIT_NOTICE_THRESHOLD (TASKS.md
  // #373). It used to show whenever a balance was known (#200), which put a
  // credit count and an Upgrade button on every screen for users with plenty
  // left. The settings menu still carries the balance at any level.
  // Not on the start screen (TASKS.md #202). `composerDock` is the SAME element
  // in both places — centred over the illustration when there are no turns, and
  // pinned to the bottom once the conversation starts — so the card followed it
  // onto a screen whose whole job is the greeting and the empty-state art. The
  // balance belongs next to work in progress, not in front of a blank slate;
  // the settings menu still carries it before anything has been asked.
  const showCreditNotice =
    creditBelowNoticeThreshold &&
    !creditNoticeDismissed &&
    !pendingQuestion &&
    !showStartScreen;

  const composerDock = (
    <div
      className={`cellix-composer-dock${
        pendingQuestion || showCreditNotice || resumableRun ? ' has-question' : ''
      }`}
    >
      {resumableRun && onResumeRun && onDismissResumableRun && (
        <ResumeRunCard
          waveIndex={resumableRun.waveIndex}
          waveTotal={resumableRun.waveTotal}
          onResume={onResumeRun}
          onDismiss={onDismissResumableRun}
          disabled={isWaitingForResponse}
        />
      )}
      {showCreditNotice && creditAccount && (
        <CreditUpgradeCard
          balance={creditAccount.availableBalance}
          isLow
          onDismiss={() => setCreditNoticeDismissed(true)}
        />
      )}
      {pendingQuestion && (
        <QuestionChoicesPanel
          key={pendingQuestion.id}
          question={pendingQuestion.question}
          options={pendingQuestion.options}
          onSelect={handleQuestionAnswer}
          disabled={isWaitingForResponse}
          selected={questionSelection}
          onSelectedChange={setQuestionSelection}
          docked
        />
      )}
      {composerInput}
    </div>
  );

  return (
    <div
      className={`cellix-panel${isDraggingFile ? ' is-dragging-file' : ''}${dragRefused ? ' is-drag-refused' : ''}`}
      onDragEnter={handlePaneDragEnter}
      onDragOver={handlePaneDragOver}
      onDragLeave={handlePaneDragLeave}
      onDrop={handlePaneDrop}
    >
      {isDraggingFile && (
        <div className="cellix-drop-overlay" aria-hidden="true">
          <div className={`cellix-drop-overlay-card${dragRefused ? ' is-refused' : ''}`}>
            <span className="cellix-drop-overlay-icons">
              {dragRefused ? (
                <AlertCircle size={32} strokeWidth={1.6} />
              ) : (
                <>
                  {(dragHasPdf || !dragHasSheet) && (
                    <span className="is-pdf">
                      <FileTypeBadge kind="pdf" size={44} />
                    </span>
                  )}
                  {(dragHasSheet || !dragHasPdf) && (
                    <span className="is-sheet">
                      <FileTypeBadge kind="sheet" size={44} />
                    </span>
                  )}
                </>
              )}
            </span>
            <strong>{dropTitle}</strong>
            <span>{dragRefused ? 'Use a PDF, XLSX, XLS or CSV file' : 'PDF, XLSX, XLS or CSV'}</span>
          </div>
        </div>
      )}
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
        creditAccount={creditAccount}
        isLowBalance={isLowCreditBalance}
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
              onRetryMissingSheets={onRetryMissingSheets}
              onRejectActions={onRejectActions}
              onResolveGstReconCollision={onResolveGstReconCollision}
              onRejectGstReconCollision={onRejectGstReconCollision}
              onJumpToGstReconRow={onJumpToGstReconRow}
              onAnswerQuestion={handleQuestionAnswer}
              dockedQuestions
              onToggleThinking={onToggleThinking}
              onAnswerComplete={onAnswerComplete}
              onFollowUp={onFollowUp}
              onRegenerate={onRegenerate}
              onRunAsAction={onRunAsAction}
              onRevertChangeSet={onRevertHistoryEntry}
            />
          ))
        )}

        {/* The standing "Last change · N cells · N min ago" bar was removed
            (TASKS.md #207): every turn that applied something already offers
            "Revert this change" in its own ↺ message-actions menu, anchored to
            the change it made, and the top bar's Change History panel still
            lists every entry. The bar restated the newest of those a third
            time, at the cost of a permanent row above the composer.
            `onRevertHistoryEntry` is still wired — it powers both survivors. */}
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
