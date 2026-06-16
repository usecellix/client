import React, { useState } from 'react';
import { ChevronRight } from 'lucide-react';
import { DiffItem } from '@/services/previewManager';

interface PreviewSummaryBarProps {
  items: DiffItem[];
  summary: string;
  onAccept: () => void;
  onReject: () => void;
  isApplying: boolean;
  showActions?: boolean;
}

export const PreviewSummaryBar: React.FC<PreviewSummaryBarProps> = ({
  items,
  summary,
  onAccept,
  onReject,
  isApplying,
  showActions = true,
}) => {
  const [expanded, setExpanded] = useState(false);
  const changeCount = items.length || 1;

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
            {changeCount} change{changeCount === 1 ? '' : 's'}
          </span>
          <span className="cellix-badge cellix-preview-badge">
            Preview
          </span>
        </button>
        <div className="cellix-preview-summary-actions">
          {showActions ? (
            <>
              <button
                type="button"
                className="cellix-preview-reject-btn"
                onClick={onReject}
                disabled={isApplying}
              >
                Reject
              </button>
              <button
                type="button"
                className="cellix-preview-accept-btn"
                onClick={onAccept}
                disabled={isApplying}
              >
                {isApplying ? 'Applying…' : 'Accept'}
              </button>
            </>
          ) : (
            <span className="cellix-preview-pending-label">Review shortly…</span>
          )}
        </div>
      </div>

      {expanded && (
        <div className="cellix-preview-summary-body">
          {summary && <p className="cellix-preview-summary-text">{summary}</p>}
          {items.length > 0 ? (
            <div className="cellix-review-table-wrap">
              <table className="cellix-review-table">
                <thead>
                  <tr>
                    <th>Cell</th>
                    <th>Action</th>
                    <th>Change</th>
                  </tr>
                </thead>
                <tbody>
                  {items.map((item, index) => (
                    <tr key={`${item.sheetName}-${item.address}-${index}`}>
                      <td>
                        <span className="cellix-cell-ref">
                          {item.sheetName}!{item.address}
                        </span>
                      </td>
                      <td>{item.actionType}</td>
                      <td>
                        <span className="cellix-after-pill">{item.description}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="cellix-preview-summary-text">Review the changes below before accepting.</p>
          )}
        </div>
      )}
    </div>
  );
};

export default PreviewSummaryBar;
