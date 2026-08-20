import React, { useState } from 'react';
import { ChevronDown, ChevronRight } from 'lucide-react';
import { ActionBlock } from '@/types/conversationTurn';
import {
  formatInternalDetailsLines,
  hasInternalDetails,
  resolveActionBlockCopy,
} from '@/utils/userFacingResponse';

export interface ActionResponseCardProps {
  block: ActionBlock;
  previewEnabled?: boolean;
  isApplying?: boolean;
  showActionButtons?: boolean;
  onAccept: () => void;
  onReject: () => void;
  /** For tests — start with details expanded. */
  defaultDetailsExpanded?: boolean;
  /**
   * Staged accept waves: set when this block's dependency (e.g. the sheet-
   * creation wave) has not been accepted yet. Disables Accept and explains why
   * instead of letting the user apply writes that would fail or land wrong.
   */
  blockedReason?: string;
}

export const ActionResponseCard: React.FC<ActionResponseCardProps> = ({
  block,
  previewEnabled = true,
  isApplying = false,
  showActionButtons = true,
  onAccept,
  onReject,
  defaultDetailsExpanded = false,
  blockedReason,
}) => {
  const [detailsOpen, setDetailsOpen] = useState(defaultDetailsExpanded);
  const summary = resolveActionBlockCopy({
    userFacingSummary: block.userFacingSummary,
    explanation: block.explanation,
    actions: block.actions,
    changes: block.changes,
  });
  const showDetails = hasInternalDetails(block.internalDetails);
  const detailLines = block.internalDetails
    ? formatInternalDetailsLines(block.internalDetails)
    : [];

  const isPending = block.proposalStatus === 'pending';
  const isAccepted = block.proposalStatus === 'accepted';
  const isRejected = block.proposalStatus === 'rejected';

  const detailsSection = showDetails ? (
    <div className="cellix-action-details" data-testid="action-details">
      <button
        type="button"
        className="cellix-thinking-toggle"
        onClick={() => setDetailsOpen((v) => !v)}
        aria-expanded={detailsOpen}
        data-testid="action-details-toggle"
      >
        {detailsOpen ? (
          <ChevronDown size={12} color="#9CA3AF" />
        ) : (
          <ChevronRight size={12} color="#9CA3AF" />
        )}
        <span>{detailsOpen ? 'Hide details' : 'Show details'}</span>
      </button>
      {detailsOpen && (
        <div
          className="cellix-thinking-body"
          data-testid="action-details-body"
          style={{ fontSize: 11.5, color: 'var(--cx-gray-500)' }}
        >
          {detailLines.map((line, i) => (
            <p key={i} style={{ margin: '4px 0' }}>
              {line}
            </p>
          ))}
        </div>
      )}
    </div>
  ) : null;

  const summaryBody = (
    <div data-testid="action-summary-default">
      {summary.contextLine && (
        <div
          className="cellix-changes-context"
          style={{ fontSize: 11.5, color: 'var(--cx-gray-500)', marginBottom: 4 }}
        >
          {summary.contextLine}
        </div>
      )}
      <div
        className="cellix-changes-summary"
        style={{ fontSize: 12.5, fontWeight: 500, color: 'var(--cx-gray-700)' }}
      >
        {summary.headline}
      </div>
      {summary.bullets && summary.bullets.length > 0 && (
        <ul style={{ margin: '6px 0 0', paddingLeft: 18, fontSize: 12, color: 'var(--cx-gray-600)' }}>
          {summary.bullets.map((b, i) => (
            <li key={i}>{b}</li>
          ))}
        </ul>
      )}
      {summary.supportingDetail && (
        <div style={{ fontSize: 11.5, color: 'var(--cx-gray-500)', marginTop: 4 }}>
          {summary.supportingDetail}
        </div>
      )}
    </div>
  );

  const acceptReject =
    isPending && showActionButtons ? (
      <div className="cellix-action-btns">
        <button
          type="button"
          className="cellix-btn-accept"
          onClick={onAccept}
          disabled={isApplying || Boolean(blockedReason)}
          title={blockedReason}
        >
          {isApplying ? 'Applying…' : 'Accept'}
        </button>
        <button
          type="button"
          className="cellix-btn-reject"
          onClick={onReject}
          disabled={isApplying}
        >
          Reject
        </button>
      </div>
    ) : null;

  const blockedNotice =
    isPending && blockedReason ? (
      <div
        className="cellix-action-blocked-notice"
        data-testid="action-blocked-notice"
        style={{ fontSize: 11.5, color: 'var(--cx-gray-500)', marginTop: 4 }}
      >
        {blockedReason}
      </div>
    ) : null;

  const hasIrreversibleActions = Boolean(block.irreversibleActionTypes?.length);
  const irreversibleNotice =
    isPending && hasIrreversibleActions ? (
      <div
        className="cellix-action-irreversible-notice"
        data-testid="action-irreversible-notice"
        style={{ fontSize: 11.5, color: 'var(--warning-amber)', marginTop: 4 }}
      >
        Not fully undoable: {block.irreversibleActionTypes!.join(', ')} — this change cannot be
        automatically reverted once applied.
      </div>
    ) : null;

  if (isAccepted) {
    return (
      <div className="cellix-changes-card cellix-block-enter">
        {summaryBody}
        <div className="cellix-changes-link" style={{ marginTop: 6 }}>
          <span style={{ fontSize: 12, color: 'var(--cx-gray-500)' }}>Applied</span>
        </div>
        {detailsSection}
      </div>
    );
  }

  if (isRejected) {
    return (
      <div className="cellix-changes-card cellix-block-enter" style={{ opacity: 0.7 }}>
        <div className="cellix-action-card-title">Changes rejected</div>
        {summaryBody}
        {detailsSection}
      </div>
    );
  }

  if (isPending && previewEnabled) {
    return (
      <div className="cellix-changes-card cellix-block-enter">
        {summaryBody}
        <div className="cellix-changes-link" style={{ marginTop: 6 }}>
          <span style={{ fontSize: 12, color: 'var(--cx-gray-500)' }}>Pending review</span>
        </div>
        {blockedNotice}
        {irreversibleNotice}
        {acceptReject}
        {detailsSection}
      </div>
    );
  }

  return (
    <div className="cellix-action-card cellix-block-enter">
      <div className="cellix-action-card-title">Cellix will make these changes:</div>
      {summaryBody}
      {blockedNotice}
      {irreversibleNotice}
      {acceptReject}
      {detailsSection}
    </div>
  );
};

export default ActionResponseCard;
