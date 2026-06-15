import React, { useCallback, useEffect, useState } from 'react';
import { History, RotateCcw } from 'lucide-react';
import { fetchChangeSetHistory, revertChangeSet } from '@/services/auditService';
import { ChangeSetSummary, formatCellValue } from '@/types/changeSet';
import { SheetAction } from '@/types/sheet-actions';

interface ChangeHistoryPanelProps {
  conversationId: string | null;
  onRevert: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
  /** When true, render the body directly (no internal toggle) for header popovers. */
  embedded?: boolean;
}

function affectedSheetsOf(entry: ChangeSetSummary): string[] {
  return Array.from(new Set(entry.changes.map((c) => c.sheet).filter(Boolean)));
}

export const ChangeHistoryPanel: React.FC<ChangeHistoryPanelProps> = ({
  conversationId,
  onRevert,
  embedded = false,
}) => {
  const [expanded, setExpanded] = useState(false);
  const [history, setHistory] = useState<ChangeSetSummary[]>([]);
  const [loading, setLoading] = useState(false);
  const [revertingId, setRevertingId] = useState<string | null>(null);

  const showBody = embedded || expanded;

  const loadHistory = useCallback(async () => {
    if (!conversationId) {
      setHistory([]);
      return;
    }
    setLoading(true);
    try {
      const items = await fetchChangeSetHistory(conversationId);
      setHistory(items);
    } catch (error) {
      console.warn('[Cellix] Failed to load change history:', error);
    } finally {
      setLoading(false);
    }
  }, [conversationId]);

  useEffect(() => {
    if (showBody) {
      void loadHistory();
    }
  }, [showBody, loadHistory]);

  const appliedCount = history.filter((h) => h.status === 'applied').length;

  if (!conversationId && !embedded) return null;

  return (
    <div className="cellix-change-history">
      {!embedded && (
        <button
          type="button"
          className="cellix-change-history-toggle"
          onClick={() => setExpanded((prev) => !prev)}
          aria-expanded={expanded}
        >
          <History size={14} />
          <span>Change history</span>
          {appliedCount > 0 && (
            <span className="cellix-badge cellix-preview-badge">{appliedCount} revertible</span>
          )}
        </button>
      )}

      {showBody && (
        <div className="cellix-change-history-body">
          {!conversationId && (
            <p className="cellix-preview-summary-text">Start a chat to track changes.</p>
          )}
          {loading && <p className="cellix-preview-summary-text">Loading…</p>}
          {!loading && conversationId && history.length === 0 && (
            <p className="cellix-preview-summary-text">No audited changes yet.</p>
          )}
          {!loading &&
            history.map((entry) => {
              const sheets = affectedSheetsOf(entry);
              return (
              <div key={entry.changeSetId} className="cellix-change-history-item">
                <div className="cellix-change-history-meta">
                  <span className="cellix-change-history-prompt">{entry.prompt}</span>
                  <span className="cellix-change-history-status">{entry.status}</span>
                  <span className="cellix-change-history-time">
                    {new Date(entry.timestamp).toLocaleString()}
                  </span>
                </div>
                <div className="cellix-change-history-summary-row">
                  {sheets.length > 0 && (
                    <span className="cellix-badge">
                      {sheets.length === 1 ? 'Sheet' : 'Sheets'}: {sheets.join(', ')}
                    </span>
                  )}
                  <span className="cellix-badge">
                    {entry.changes.length} cell{entry.changes.length === 1 ? '' : 's'} changed
                  </span>
                </div>
                <div className="cellix-change-history-cells">
                  {entry.changes.slice(0, 3).map((c, i) => (
                    <span key={`${entry.changeSetId}-${i}`} className="cellix-change-history-chip">
                      {c.sheet}!{c.cell}: {formatCellValue(c.before)} → {formatCellValue(c.after)}
                    </span>
                  ))}
                  {entry.changes.length > 3 && (
                    <span className="cellix-change-history-chip">
                      +{entry.changes.length - 3} more
                    </span>
                  )}
                </div>
                {entry.status === 'applied' && (
                  <button
                    type="button"
                    className="cellix-change-history-revert"
                    disabled={revertingId === entry.changeSetId}
                    onClick={async () => {
                      setRevertingId(entry.changeSetId);
                      try {
                        const result = await revertChangeSet(entry.changeSetId);
                        await onRevert(entry.changeSetId, result.inverseActions);
                        await loadHistory();
                      } catch (error) {
                        console.error('[Cellix] Revert failed:', error);
                      } finally {
                        setRevertingId(null);
                      }
                    }}
                  >
                    <RotateCcw size={12} />
                    {revertingId === entry.changeSetId ? 'Reverting…' : 'Revert'}
                  </button>
                )}
              </div>
              );
            })}
        </div>
      )}
    </div>
  );
};

export default ChangeHistoryPanel;
