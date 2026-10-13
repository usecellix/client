import React from 'react';
import { GstReconCollisionBlock } from '@/types/conversationTurn';

export interface GstReconCollisionCardProps {
  block: GstReconCollisionBlock;
  disabled?: boolean;
  onOverwrite: () => void;
  onCreateNew: () => void;
  onReject: () => void;
}

/**
 * ActionResponseCard-style choice card for a sheet-name collision: Overwrite,
 * Create new, or Reject, as buttons — clicking Overwrite/Create new directly
 * resolves the held-in-memory reconciliation (see `resolveGstReconSheetCollision`),
 * no chat-text round trip or NLU matching involved. Reject discards the pending
 * collision and the reconciliation result outright, same as the plain Accept/Reject
 * actions card's Reject — nothing is written to the workbook either way.
 */
export const GstReconCollisionCard: React.FC<GstReconCollisionCardProps> = ({
  block,
  disabled = false,
  onOverwrite,
  onCreateNew,
  onReject,
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
        <button
          type="button"
          className="cellix-btn-reject"
          onClick={onReject}
          disabled={disabled || isResolving}
          data-testid="gst-recon-collision-reject"
        >
          Reject
        </button>
      </div>
    </div>
  );
};

export default GstReconCollisionCard;
