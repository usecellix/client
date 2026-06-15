import React, { useState } from 'react';
import { ChevronRight } from 'lucide-react';
import { CellChange, formatCellValue } from '@/types/changeSet';

interface DiffPreviewPanelProps {
  changes: CellChange[];
  summary: string;
  changeSetId?: string;
  onAccept: () => void;
  onReject: () => void;
  isApplying: boolean;
}

export const DiffPreviewPanel: React.FC<DiffPreviewPanelProps> = ({
  changes,
  summary,
  changeSetId,
  onAccept,
  onReject,
  isApplying,
}) => {
  const [expanded, setExpanded] = useState(true);
  const [localApplying, setLocalApplying] = useState(false);
  const changeCount = changes.length || 1;
  const applying = isApplying || localApplying;

  const handleAccept = () => {
    if (applying) return;
    setLocalApplying(true);
    void Promise.resolve(onAccept()).finally(() => setLocalApplying(false));
  };

  return (
    <div className="cellix-preview-summary">
      <div className="cellix-preview-summary-bar">
        <button
          type="button"
          className="cellix-preview-summary-toggle"
          onClick={() => setExpanded((prev) => !prev)}
          aria-expanded={expanded}
        >
          <ChevronRight
            size={14}
            className={`cellix-preview-chevron${expanded ? ' expanded' : ''}`}
          />
          <span className="cellix-preview-count">
            {changeCount} cell change{changeCount === 1 ? '' : 's'}
          </span>
          <span className="cellix-badge cellix-preview-badge">
            {changeSetId ? 'Audited diff' : 'Preview'}
          </span>
        </button>
        <div className="cellix-preview-summary-actions">
          <button
            type="button"
            className="cellix-preview-reject-btn"
            onClick={onReject}
            disabled={applying}
          >
            Reject
          </button>
          <button
            type="button"
            className="cellix-preview-accept-btn"
            onClick={handleAccept}
            disabled={applying}
          >
            {applying ? 'Applying…' : 'Apply'}
          </button>
        </div>
      </div>

      {expanded && (
        <div className="cellix-preview-summary-body">
          {summary && <p className="cellix-preview-summary-text">{summary}</p>}
          {changes.length > 0 ? (
            <div className="cellix-review-table-wrap">
              <table className="cellix-review-table">
                <thead>
                  <tr>
                    <th>Cell</th>
                    <th>Before</th>
                    <th>After</th>
                    <th>Formula</th>
                    <th>Hardcoded</th>
                  </tr>
                </thead>
                <tbody>
                  {changes.map((item, index) => (
                    <tr key={`${item.sheet}-${item.cell}-${index}`}>
                      <td>
                        <span className="cellix-cell-ref">
                          {item.sheet}!{item.cell}
                        </span>
                      </td>
                      <td className="cellix-diff-before">{formatCellValue(item.before)}</td>
                      <td className="cellix-diff-after">{formatCellValue(item.after)}</td>
                      <td>{item.formula ?? '—'}</td>
                      <td>{item.isHardcoded ? 'Yes' : 'No'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="cellix-preview-summary-text">Review the changes below before applying.</p>
          )}
        </div>
      )}
    </div>
  );
};

export default DiffPreviewPanel;
