import React, { useCallback, useEffect, useState } from 'react';
import { Flag, RotateCcw, Plus } from 'lucide-react';
import { fetchCheckpoints, createCheckpoint, restoreCheckpoint } from '@/services/checkpointService';
import { CheckpointSummary, RestoreResult } from '@/types/checkpoint';

interface CheckpointPanelProps {
  workbookId: string | undefined;
  conversationId: string | null;
  onRestore: (result: RestoreResult) => Promise<void>;
  /** When true, render the body directly (no internal toggle) for header popovers. */
  embedded?: boolean;
}

export const CheckpointPanel: React.FC<CheckpointPanelProps> = ({
  workbookId,
  conversationId,
  onRestore,
  embedded = false,
}) => {
  const [expanded, setExpanded] = useState(false);
  const [checkpoints, setCheckpoints] = useState<CheckpointSummary[]>([]);
  const [loading, setLoading] = useState(false);
  const [creating, setCreating] = useState(false);
  const [restoringId, setRestoringId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const showBody = embedded || expanded;

  const loadCheckpoints = useCallback(async () => {
    if (!workbookId) {
      setCheckpoints([]);
      return;
    }
    setLoading(true);
    try {
      const items = await fetchCheckpoints(workbookId);
      setCheckpoints(items);
    } catch (err) {
      console.warn('[Cellix] Failed to load checkpoints:', err);
    } finally {
      setLoading(false);
    }
  }, [workbookId]);

  useEffect(() => {
    if (showBody) {
      void loadCheckpoints();
    }
  }, [showBody, loadCheckpoints]);

  const handleCreate = useCallback(async () => {
    if (!workbookId || !conversationId) return;
    setCreating(true);
    setError(null);
    try {
      await createCheckpoint(workbookId, conversationId);
      await loadCheckpoints();
    } catch (err) {
      console.error('[Cellix] Checkpoint creation failed:', err);
      setError(err instanceof Error ? err.message : 'Failed to create checkpoint');
    } finally {
      setCreating(false);
    }
  }, [workbookId, conversationId, loadCheckpoints]);

  const handleRestore = useCallback(
    async (checkpointId: string) => {
      setRestoringId(checkpointId);
      setError(null);
      try {
        const result = await restoreCheckpoint(checkpointId);
        await onRestore(result);
        await loadCheckpoints();
      } catch (err) {
        // TASKS.md #29 — the backend fails closed (422) and names the specific
        // blocking change set rather than partially restoring; surface that
        // message verbatim rather than a generic "something went wrong."
        console.error('[Cellix] Restore failed:', err);
        setError(err instanceof Error ? err.message : 'Restore failed');
      } finally {
        setRestoringId(null);
      }
    },
    [onRestore, loadCheckpoints],
  );

  const activeCount = checkpoints.filter((c) => c.status === 'active').length;

  if (!workbookId && !embedded) return null;

  return (
    <div className="cellix-checkpoint-panel">
      {!embedded && (
        <button
          type="button"
          className="cellix-checkpoint-toggle"
          onClick={() => setExpanded((prev) => !prev)}
          aria-expanded={expanded}
        >
          <Flag size={14} />
          <span>Checkpoints</span>
          {activeCount > 0 && (
            <span className="cellix-badge cellix-preview-badge">{activeCount}</span>
          )}
        </button>
      )}

      {showBody && (
        <div className="cellix-checkpoint-body">
          {!workbookId && (
            <p className="cellix-preview-summary-text">Open a workbook to use checkpoints.</p>
          )}
          {workbookId && (
            <button
              type="button"
              className="cellix-checkpoint-create"
              disabled={creating || !conversationId}
              onClick={() => void handleCreate()}
            >
              <Plus size={12} />
              {creating ? 'Creating…' : 'Create checkpoint'}
            </button>
          )}
          {error && <p className="cellix-checkpoint-error">{error}</p>}
          {loading && <p className="cellix-preview-summary-text">Loading…</p>}
          {!loading && workbookId && checkpoints.length === 0 && (
            <p className="cellix-preview-summary-text">No checkpoints yet for this workbook.</p>
          )}
          {!loading &&
            checkpoints.map((cp) => (
              <div key={cp.checkpointId} className="cellix-checkpoint-item">
                <div className="cellix-checkpoint-meta">
                  <span className="cellix-checkpoint-label">{cp.label}</span>
                  <span className={`cellix-checkpoint-trigger cellix-checkpoint-trigger-${cp.trigger}`}>
                    {cp.trigger}
                  </span>
                  <span className="cellix-checkpoint-status">{cp.status}</span>
                </div>
                <div className="cellix-checkpoint-time">
                  {new Date(cp.createdAt).toLocaleString()}
                </div>
                {cp.status === 'active' && (
                  <button
                    type="button"
                    className="cellix-checkpoint-restore"
                    disabled={restoringId === cp.checkpointId}
                    onClick={() => void handleRestore(cp.checkpointId)}
                  >
                    <RotateCcw size={12} />
                    {restoringId === cp.checkpointId ? 'Restoring…' : 'Restore'}
                  </button>
                )}
              </div>
            ))}
        </div>
      )}
    </div>
  );
};

export default CheckpointPanel;
