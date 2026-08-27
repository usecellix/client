import React, { useCallback, useEffect, useState } from 'react';
import { RotateCcw } from 'lucide-react';
import { fetchChangeSetHistory, revertChangeSet } from '@/services/auditService';
import { ChangeSetSummary } from '@/types/changeSet';
import { SheetAction } from '@/types/sheet-actions';

/**
 * Revert affordance for the most recent applied change, rendered inline at the
 * end of the conversation rather than behind the composer's history icon.
 *
 * "Undo the thing that just happened" is the overwhelmingly common case, and it
 * previously required opening a panel and finding the top entry. The full
 * per-entry history still lives in ChangeHistoryPanel — this is a shortcut to
 * its first row, not a replacement for it.
 */
interface LastChangeRevertProps {
  conversationId: string | null;
  onRevert: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
  /** Bump to re-fetch — e.g. after an accept applies a new change set. */
  refreshKey?: string | number;
}

export const LastChangeRevert: React.FC<LastChangeRevertProps> = ({
  conversationId,
  onRevert,
  refreshKey,
}) => {
  const [entry, setEntry] = useState<ChangeSetSummary | null>(null);
  const [reverting, setReverting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadLatest = useCallback(async () => {
    if (!conversationId) {
      setEntry(null);
      return;
    }
    try {
      const history = await fetchChangeSetHistory(conversationId);
      // Only an applied change can be reverted; anything already reverted or
      // still pending must not offer the button.
      setEntry(history.find((e) => e.status === 'applied') ?? null);
    } catch (err) {
      console.error('[Cellix] Could not load latest change for revert:', err);
      setEntry(null);
    }
  }, [conversationId]);

  useEffect(() => {
    void loadLatest();
  }, [loadLatest, refreshKey]);

  if (!entry) return null;

  const cellCount = entry.changes.length;
  const sheets = Array.from(new Set(entry.changes.map((c) => c.sheet).filter(Boolean)));
  const scope =
    sheets.length === 1 ? sheets[0] : sheets.length > 1 ? `${sheets.length} sheets` : '';

  return (
    <div className="cellix-last-change-revert" data-testid="last-change-revert">
      <span className="cellix-last-change-revert-label">
        Last change{scope ? ` on ${scope}` : ''}
        {cellCount ? ` · ${cellCount} cell${cellCount === 1 ? '' : 's'}` : ''}
      </span>
      <button
        type="button"
        className="cellix-last-change-revert-btn"
        disabled={reverting}
        aria-label="Revert the last applied change"
        onClick={async () => {
          setReverting(true);
          setError(null);
          try {
            const result = await revertChangeSet(entry.changeSetId);
            await onRevert(entry.changeSetId, result.inverseActions);
            await loadLatest();
          } catch (err) {
            console.error('[Cellix] Revert failed:', err);
            // Surface the failure instead of leaving the row looking unchanged —
            // a silent no-op here is the same class of bug as TASKS.md #80.
            setError('Revert failed. Open change history to retry.');
          } finally {
            setReverting(false);
          }
        }}
      >
        <RotateCcw size={12} />
        {reverting ? 'Reverting…' : 'Revert'}
      </button>
      {error && <span className="cellix-last-change-revert-error">{error}</span>}
    </div>
  );
};

export default LastChangeRevert;
