import React from 'react';
import { ExternalLink, AlertTriangle } from 'lucide-react';
import {
  DomainException,
  SourceRef,
  formatSourceRefLabel,
} from '@/types/changeSet';

export interface SourcePreviewProps {
  sourceRefs: SourceRef[];
  exceptionFlags?: DomainException[];
  onJumpToSource: (ref: SourceRef) => void;
  onClose?: () => void;
}

export const SourcePreview: React.FC<SourcePreviewProps> = ({
  sourceRefs,
  exceptionFlags = [],
  onJumpToSource,
  onClose,
}) => {
  return (
    <div className="cellix-source-preview" data-testid="source-preview">
      <div className="cellix-source-preview-header">
        <span className="cellix-source-preview-title">Source</span>
        {onClose && (
          <button type="button" className="cellix-source-preview-close" onClick={onClose}>
            Close
          </button>
        )}
      </div>

      {exceptionFlags.length > 0 && (
        <ul className="cellix-source-exceptions" data-testid="source-exceptions">
          {exceptionFlags.map((flag, index) => (
            <li
              key={`${flag.code}-${index}`}
              className={`cellix-source-exception cellix-source-exception-${flag.severity}`}
            >
              <AlertTriangle size={12} />
              <span>
                <strong>{flag.code}</strong> — {flag.message}
              </span>
            </li>
          ))}
        </ul>
      )}

      {sourceRefs.length === 0 ? (
        <p className="cellix-source-preview-empty">No source citations for this change.</p>
      ) : (
        <ul className="cellix-source-ref-list">
          {sourceRefs.map((ref, index) => (
            <li key={`${ref.documentType}-${ref.rowOrLine}-${index}`}>
              <button
                type="button"
                className="cellix-source-ref-btn"
                onClick={() => onJumpToSource(ref)}
                title={
                  ref.documentType === 'workbook'
                    ? 'Select this range in Excel'
                    : 'Open source document location'
                }
              >
                <ExternalLink size={12} />
                <span>{formatSourceRefLabel(ref)}</span>
                <span className="cellix-source-ref-type">{ref.documentType}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};

export default SourcePreview;
