import React, { useEffect, useRef, useState } from 'react';
import { ChevronRight } from 'lucide-react';
import { ActionBlock } from '@/types/conversationTurn';
import { resolveActionBlockCopy } from '@/utils/userFacingResponse';
import { describeSheetActions } from '@/utils/describeSheetAction';

function normalizeCopy(text: string): string {
  return text
    .trim()
    .toLowerCase()
    .replace(/[.…]+$/g, '')
    .replace(/\s+/g, ' ');
}

export interface ActionResponseCardProps {
  block: ActionBlock;
  previewEnabled?: boolean;
  isApplying?: boolean;
  showActionButtons?: boolean;
  onAccept: () => void;
  onReject: () => void;
  /**
   * Accept this step and every remaining one, in order — TASKS.md #160.
   * Omitted when this is not part of a staged build, or is the last step.
   */
  onAcceptAll?: () => void;
  /** For tests — start with details expanded. */
  defaultDetailsExpanded?: boolean;
  /**
   * Staged accept waves: set when this block's dependency (e.g. the sheet-
   * creation wave) has not been accepted yet. Disables Accept and explains why
   * instead of letting the user apply writes that would fail or land wrong.
   */
  blockedReason?: string;
  /**
   * Answer text already shown above this card. When it matches the action
   * headline, omit the repeated line so the user does not see the same sentence twice.
   */
  priorAnswerText?: string;
}

export const ActionResponseCard: React.FC<ActionResponseCardProps> = ({
  block,
  previewEnabled = true,
  isApplying = false,
  showActionButtons = true,
  onAccept,
  onReject,
  onAcceptAll,
  defaultDetailsExpanded = false,
  blockedReason,
  priorAnswerText,
}) => {
  const [detailsOpen, setDetailsOpen] = useState(defaultDetailsExpanded);
  const detailsBodyRef = useRef<HTMLDivElement>(null);

  // Local toggle state doesn't touch the conversation's turns array, so the
  // panel's sticky-to-bottom scroll never sees it expand — bring the newly
  // revealed body into view ourselves instead of leaving it hidden behind
  // the composer.
  useEffect(() => {
    if (detailsOpen) {
      detailsBodyRef.current?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  }, [detailsOpen]);
  const summary = resolveActionBlockCopy({
    userFacingSummary: block.userFacingSummary,
    explanation: block.explanation,
    actions: block.actions,
    changes: block.changes,
  });
  const hideDuplicateHeadline = Boolean(
    priorAnswerText &&
      summary.headline &&
      normalizeCopy(priorAnswerText) === normalizeCopy(summary.headline),
  );
  const showDetails = block.actions.length > 0;
  const detailLines = showDetails ? describeSheetActions(block.actions) : [];

  const isPending = block.proposalStatus === 'pending';
  const isAccepted = block.proposalStatus === 'accepted';
  const isRejected = block.proposalStatus === 'rejected';

  const detailsToggle = showDetails ? (
    <button
      type="button"
      className="cellix-details-toggle"
      onClick={() => setDetailsOpen((v) => !v)}
      aria-expanded={detailsOpen}
      data-testid="action-details-toggle"
    >
      <ChevronRight
        size={11}
        className={`cellix-details-chevron${detailsOpen ? ' is-open' : ''}`}
      />
      <span>{detailsOpen ? 'Hide details' : 'Show details'}</span>
    </button>
  ) : null;

  const detailsBody = showDetails ? (
    <div className="cellix-action-details" data-testid="action-details">
      {detailsOpen && (
        <div className="cellix-details-body" data-testid="action-details-body" ref={detailsBodyRef}>
          {detailLines.map((line, i) => (
            <p key={i}>{line}</p>
          ))}
        </div>
      )}
    </div>
  ) : null;

  /** Meta line: the one muted row under the headline (scope + size). */
  const metaParts = [summary.contextLine, summary.supportingDetail].filter(Boolean) as string[];

  /**
   * Position within a staged build. This is the piece TASKS.md #141's version
   * lacked: without it, accepting one card and stopping left a half-built
   * workbook that looked finished. Stating "Step 2 of 5" makes the remaining
   * work visible rather than inferred. TASKS.md #160.
   */
  const isStaged =
    typeof block.stepIndex === 'number' &&
    typeof block.stepTotal === 'number' &&
    block.stepTotal > 1;
  const stepBadge = isStaged ? (
    <div className="cellix-step-badge" data-testid="action-step-badge">
      <span className="cellix-step-count">
        Step {block.stepIndex} of {block.stepTotal}
      </span>
      {block.stepLabel && <span className="cellix-step-label">{block.stepLabel}</span>}
      {isPending && block.stepIndex! < block.stepTotal! && (
        <span className="cellix-step-remaining">
          {block.stepTotal! - block.stepIndex!} more step
          {block.stepTotal! - block.stepIndex! === 1 ? '' : 's'} after this
        </span>
      )}
    </div>
  ) : null;

  const hasSummaryContent = Boolean(
    stepBadge ||
      (!hideDuplicateHeadline && summary.headline) ||
      (summary.bullets && summary.bullets.length > 0) ||
      metaParts.length > 0,
  );

  const summaryBody = hasSummaryContent ? (
    <div data-testid="action-summary-default">
      {stepBadge}
      {!hideDuplicateHeadline && summary.headline && (
        <div className="cellix-changes-summary">{summary.headline}</div>
      )}
      {summary.bullets && summary.bullets.length > 0 && (
        <ul className="cellix-changes-bullets">
          {summary.bullets.map((b, i) => (
            <li key={i}>{b}</li>
          ))}
        </ul>
      )}
      {metaParts.length > 0 && (
        <div className="cellix-changes-meta">{metaParts.join(' · ')}</div>
      )}
    </div>
  ) : null;

  const compactClass = hasSummaryContent ? '' : ' is-compact';

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
        {onAcceptAll && (
          <button
            type="button"
            className="cellix-btn-accept-all"
            onClick={onAcceptAll}
            disabled={isApplying || Boolean(blockedReason)}
            title="Apply this step and every remaining step, in order"
          >
            Accept All
          </button>
        )}
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
      <div className="cellix-action-notice" data-testid="action-blocked-notice">
        {blockedReason}
      </div>
    ) : null;

  const hasIrreversibleActions = Boolean(block.irreversibleActionTypes?.length);
  const irreversibleNotice =
    isPending && hasIrreversibleActions ? (
      <div
        className="cellix-action-notice is-warning"
        data-testid="action-irreversible-notice"
      >
        Can’t be undone automatically: {block.irreversibleActionTypes!.join(', ')}
      </div>
    ) : null;

  /**
   * One footer row carries both the details affordance and the decision, so the
   * card costs a single line of height instead of three stacked ones.
   */
  const footer = (statusNode: React.ReactNode) =>
    detailsToggle || acceptReject || statusNode ? (
      <div className="cellix-action-footer">
        <div className="cellix-action-footer-left">
          {detailsToggle}
          {statusNode}
        </div>
        {acceptReject}
      </div>
    ) : null;

  if (isAccepted) {
    return (
      <div className={`cellix-changes-card cellix-block-enter is-applied${compactClass}`}>
        {summaryBody}
        {footer(<span className="cellix-action-status is-applied">Applied</span>)}
        {detailsBody}
      </div>
    );
  }

  if (isRejected) {
    return (
      <div className={`cellix-changes-card cellix-block-enter is-rejected${compactClass}`}>
        {summaryBody}
        {footer(<span className="cellix-action-status">Rejected</span>)}
        {detailsBody}
      </div>
    );
  }

  if (isPending && previewEnabled) {
    return (
      <div className={`cellix-changes-card cellix-block-enter is-pending${compactClass}`}>
        {summaryBody}
        {blockedNotice}
        {irreversibleNotice}
        {footer(
          showActionButtons ? null : <span className="cellix-action-status">Pending review</span>,
        )}
        {detailsBody}
      </div>
    );
  }

  return (
    <div className={`cellix-action-card cellix-block-enter${compactClass}`}>
      <div className="cellix-action-card-title">Cellix will make these changes</div>
      {summaryBody}
      {blockedNotice}
      {irreversibleNotice}
      {footer(null)}
      {detailsBody}
    </div>
  );
};

export default ActionResponseCard;
