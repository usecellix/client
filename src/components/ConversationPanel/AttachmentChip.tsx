import React, { useCallback, useEffect, useRef, useState } from 'react';
import { AlertCircle, ArrowRight, FileSpreadsheet, FileText, Loader2, Lock, X } from 'lucide-react';
import { AttachmentError, attachmentExtension } from '@/services/attachments/attachmentError';
import { decodeAttachment } from '@/services/attachments/decodeAttachment';
import { RawTable } from '@/types/rawTable';

interface AttachmentChipProps {
  file: File;
  /**
   * True while the assistant is busy. The file is still read, but the import
   * waits until this clears, so it never starts on top of a running turn.
   */
  disabled?: boolean;
  onRemove: () => void;
  /**
   * Hands over the decoded file. Returns false when the import cannot start
   * yet; the chip then keeps the file and offers it again. Omit to show the
   * chip without importing anything.
   */
  onImport?: (rawTable: RawTable) => boolean;
  /**
   * Bumped by the composer each time the user presses Send. A file that has
   * been read waits for this before it imports, so the user can type a
   * question to go with it first. Typing a PDF's password only opens the file.
   */
  importRequest?: number;
  /** Tells the composer whether this file has been read and is waiting for Send. */
  onReadyChange?: (ready: boolean) => void;
}

/** `ready`: read (and unlocked if it had a password), waiting for the user to press Send. */
type Phase = 'idle' | 'reading' | 'password' | 'ready' | 'waiting' | 'error';

/** Pause after the last keystroke before a typed password is tried on its own. */
const PASSWORD_AUTO_TRY_DELAY_MS = 600;

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * One attached file, as a small capsule. It is read as soon as it is attached,
 * here in the task pane (the file itself is not uploaded), and then waits in
 * the message box like an attachment in any chat: it goes when the user presses
 * Send, with whatever they typed alongside it. A password-protected PDF shows a
 * short password field beside the capsule first. Typing the right password unlocks
 * the file by itself (Enter also works). The password is used for that one read and then dropped.
 * Root/ATTACHMENT_EXTRACTION_PLAN.md §6.5.
 */
const AttachmentChip: React.FC<AttachmentChipProps> = ({
  file,
  disabled = false,
  onRemove,
  onImport,
  importRequest = 0,
  onReadyChange,
}) => {
  const [phase, setPhase] = useState<Phase>(onImport ? 'reading' : 'idle');
  const [password, setPassword] = useState('');
  const [wrongPassword, setWrongPassword] = useState<string | null>(null);
  const [failure, setFailure] = useState<string | null>(null);
  const [decoded, setDecoded] = useState<RawTable | null>(null);
  // Set true when Send is pressed while this file is waiting.
  const [released, setReleased] = useState(false);
  const seenRequestRef = useRef(importRequest);
  // Set false on unmount, so removing a chip mid-read discards the result.
  const mountedRef = useRef(true);
  const startedRef = useRef(false);

  const read = useCallback(
    async (withPassword?: string) => {
      setPhase('reading');
      setFailure(null);
      try {
        const rawTable = await decodeAttachment(file, withPassword);
        if (!mountedRef.current) return;
        setPassword('');
        setWrongPassword(null);
        setDecoded(rawTable);
        setReleased(false);
        setPhase('ready');
      } catch (error) {
        if (!mountedRef.current) return;
        if (error instanceof AttachmentError && error.code === 'password_required') {
          setPhase('password');
        } else if (error instanceof AttachmentError && error.code === 'password_incorrect') {
          setPhase('password');
          setPassword('');
          setWrongPassword(error.message);
        } else {
          setPhase('error');
          setFailure(error instanceof AttachmentError ? error.message : 'This file could not be read.');
        }
      }
    },
    [file],
  );

  // While the user types a password, try it quietly once they pause. A wrong
  // guess says nothing and keeps the field as it is, so the right password
  // unlocks the file without pressing Enter. Enter still works and still reports
  // a wrong password. Each attempt is a full open of the file, hence the pause.
  useEffect(() => {
    if (phase !== 'password' || !password) return;
    const attempted = password;
    const timer = window.setTimeout(async () => {
      try {
        const rawTable = await decodeAttachment(file, attempted);
        if (!mountedRef.current) return;
        setPassword('');
        setWrongPassword(null);
        setDecoded(rawTable);
        setReleased(false);
        setPhase('ready');
      } catch (error) {
        if (!mountedRef.current) return;
        if (error instanceof AttachmentError && error.code === 'password_incorrect') return;
        setPhase('error');
        setFailure(error instanceof AttachmentError ? error.message : 'This file could not be read.');
      }
    }, PASSWORD_AUTO_TRY_DELAY_MS);
    return () => window.clearTimeout(timer);
  }, [password, phase, file]);

  // Start reading the moment the file is attached. The ref keeps a development
  // double-mount from reading the same file twice.
  useEffect(() => {
    mountedRef.current = true;
    if (onImport && !startedRef.current) {
      startedRef.current = true;
      void read();
    }
    return () => {
      mountedRef.current = false;
    };
    // Runs once per chip: a chip is for one file and is never reused.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Send was pressed: a file that is waiting may now go. Only a press made
  // while it is waiting counts, so an earlier Send cannot release a later file.
  useEffect(() => {
    if (importRequest === seenRequestRef.current) return;
    seenRequestRef.current = importRequest;
    if (phase === 'ready') {
      setReleased(true);
      setPhase('waiting');
    }
  }, [importRequest, phase]);

  const isReady = phase === 'ready';
  useEffect(() => {
    onReadyChange?.(isReady);
    return () => {
      if (isReady) onReadyChange?.(false);
    };
    // The callback is recreated on every composer render; only the state matters here.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isReady]);

  // Hand the decoded file over as soon as it is released and the assistant is free.
  useEffect(() => {
    if (!decoded || !released || disabled || !onImport) return;
    if (onImport(decoded)) setDecoded(null);
  }, [decoded, released, disabled, onImport]);

  const extension = attachmentExtension(file.name);
  const details = [extension?.toUpperCase(), formatSize(file.size)].filter(Boolean).join(' · ');
  const status =
    phase === 'reading'
      ? 'Reading…'
      : phase === 'waiting'
        ? 'Importing…'
        : phase === 'password'
          ? 'Password protected'
          : phase === 'ready'
            ? 'Ready. Press Enter to import.'
            : phase === 'error'
              ? 'Could not be read'
              : details;

  // The icon carries the state, so the capsule stays as small as the name. A
  // file that is ready shows its own type (red document for a PDF, green sheet
  // for Excel or CSV): once read it is just a file waiting to be sent.
  const icon =
    phase === 'reading' || phase === 'waiting' ? (
      <Loader2 size={13} className="cellix-attachment-spinner" />
    ) : phase === 'password' ? (
      <Lock size={13} />
    ) : phase === 'error' ? (
      <AlertCircle size={13} />
    ) : extension === 'pdf' ? (
      <FileText size={13} />
    ) : (
      <FileSpreadsheet size={13} />
    );

  return (
    <div className={`cellix-attachment is-${phase}`}>
      <div className="cellix-attachment-chip" title={`${file.name} (${details}) — ${status}`}>
        <span className={`cellix-attachment-icon${extension === 'pdf' ? ' is-pdf' : ' is-sheet'}`} aria-hidden="true">
          {icon}
        </span>
        <span className="cellix-attachment-name">{file.name}</span>
        <span className="cellix-sr-only" role="status">
          {status}
        </span>
        {phase === 'error' && onImport && (
          <button type="button" className="cellix-attachment-retry" onClick={() => void read()}>
            Retry
          </button>
        )}
        <button
          type="button"
          className="cellix-attachment-remove"
          onClick={onRemove}
          aria-label={`Remove ${file.name}`}
        >
          <X size={12} />
        </button>
      </div>
      {phase === 'password' && (
        <form
          className={`cellix-attachment-password${wrongPassword ? ' is-wrong' : ''}`}
          onSubmit={(event) => {
            event.preventDefault();
            if (password) void read(password);
          }}
        >
          <input
            type="password"
            value={password}
            onChange={(event) => {
              setPassword(event.target.value);
              setWrongPassword(null);
            }}
            placeholder={wrongPassword ? 'Wrong password, try again' : 'PDF password'}
            title="This PDF is password protected. Type its password; it unlocks on its own, or press Enter."
            aria-label={`Password for ${file.name}`}
            aria-invalid={wrongPassword ? true : undefined}
            autoComplete="off"
            autoFocus
          />
          <button type="submit" disabled={!password} aria-label="Unlock" title="Unlock (Enter)">
            <ArrowRight size={12} />
          </button>
          {wrongPassword && (
            <span className="cellix-sr-only" role="alert">
              {wrongPassword}
            </span>
          )}
        </form>
      )}
      {failure && (
        <div className="cellix-attachment-note is-error" role="alert">
          {failure}
        </div>
      )}
    </div>
  );
};

export default AttachmentChip;
