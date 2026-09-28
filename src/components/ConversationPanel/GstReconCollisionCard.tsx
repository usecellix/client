import React from 'react';
import { GstReconCollisionBlock } from '@/types/conversationTurn';

export interface GstReconCollisionCardProps {
  block: GstReconCollisionBlock;
  disabled?: boolean;
  onOverwrite: () => void;
  onCreateNew: () => void;
}

/**
 * ActionResponseCard-style choice card for a sheet-name collision: Overwrite
 * or Create new, as buttons — clicking one directly resolves the held-in-
 * memory reconciliation (see `resolveGstReconSheetCollision`), no chat-text
 * round trip or NLU matching involved.
 */
export const GstReconCollisionCard: React.FC<GstReconCollisionCardProps> = ({
  block,
  disabled = false,
  onOverwrite,
  onCreateNew,
}) => {
  const isResolving = Boolean(block.resolving);

  return (
    <div
      className="cellix-changes-card cellix-block-enter is-pending"
      data-testid="gst-recon-collision-card"
    >
      <div className="cellix-changes-summary">
        A sheet named "{block.sheetName}" already exists — overwrite it, or create a new one?
      </div>
      <div className="cellix-action-btns">
        <button
          type="button"
          className="cellix-btn-accept"
          onClick={() => onOverwrite()}
          disabled={disabled || isResolving}
          data-testid="gst-recon-collision-overwrite"
        >
          {isResolving ? 'Working…' : 'Overwrite'}
        </button>
        <button
          type="button"
          className="cellix-btn-reject"
          onClick={onCreateNew}
          disabled={disabled || isResolving}
          data-testid="gst-recon-collision-create-new"
        >
          {isResolving ? 'Working…' : 'Create new'}
        </button>
      </div>
    </div>
  );
};

export default GstReconCollisionCard;
