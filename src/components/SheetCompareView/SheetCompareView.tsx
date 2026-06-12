import React from 'react';
import { CompareResult } from '@/types/cellix.types';

export type { CompareResult };

interface SheetCompareViewProps {
  result: CompareResult | null;
  isLoading: boolean;
  onClose: () => void;
}

export const SheetCompareView: React.FC<SheetCompareViewProps> = ({
  result,
  isLoading,
  onClose,
}) => {
  if (isLoading) {
    return (
      <div className="cellix-compare-loading">
        <div className="cellix-compare-spinner">⟳</div>
        <p>Comparing sheets…</p>
      </div>
    );
  }

  if (!result) return null;

  return (
    <div className="cellix-compare-view">
      <div className="cellix-compare-header">
        <div>
          <h3 className="cellix-compare-title">
            {result.sheetA} ⇄ {result.sheetB}
          </h3>
          <p className="cellix-compare-summary">{result.summary}</p>
        </div>
        <button type="button" onClick={onClose} className="cellix-compare-close">
          ✕
        </button>
      </div>

      <div className="cellix-compare-stats">
        <div className="cellix-compare-stat">
          <div className="cellix-compare-stat-value added">{result.addedInB.length}</div>
          <div className="cellix-compare-stat-label">Added in B</div>
        </div>
        <div className="cellix-compare-stat">
          <div className="cellix-compare-stat-value removed">{result.removedInB.length}</div>
          <div className="cellix-compare-stat-label">Removed in B</div>
        </div>
        <div className="cellix-compare-stat">
          <div className="cellix-compare-stat-value modified">{result.modifiedCells.length}</div>
          <div className="cellix-compare-stat-label">Modified</div>
        </div>
      </div>

      <div className="cellix-compare-diff-list">
        {result.modifiedCells.slice(0, 20).map((cell) => (
          <div key={`${cell.address}-${cell.valueA}-${cell.valueB}`} className="cellix-compare-diff-row">
            <span className="cellix-compare-cell-address">{cell.address}</span>
            <span className="cellix-compare-cell-before">{cell.valueA}</span>
            <span className="cellix-compare-cell-arrow">→</span>
            <span className="cellix-compare-cell-after">{cell.valueB}</span>
          </div>
        ))}
        {result.modifiedCells.length > 20 && (
          <p className="cellix-compare-more">
            +{result.modifiedCells.length - 20} more differences
          </p>
        )}
        {result.modifiedCells.length === 0 && (
          <p className="cellix-compare-empty">No cell modifications found.</p>
        )}
      </div>
    </div>
  );
};

export default SheetCompareView;
