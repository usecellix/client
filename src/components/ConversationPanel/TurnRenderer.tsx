import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Check, Pencil, RefreshCw, RotateCcw, X } from 'lucide-react';
import {
  ActionBlock,
  AnswerBlock,
  ConversationTurn,
  PlanBlock,
  TurnBlock,
} from '@/types/conversationTurn';
import { generateSuggestedFollowUps } from '@/utils/suggestedFollowUps';
import { shortenActionPreviewCopy } from '@/utils/actionPreviewCopy';
import { describeBlockedReason } from '@/utils/actionWaveGating';
import { isTurnPresentationComplete } from '@/utils/turnPresentation';
import { revertChangeSet } from '@/services/auditService';
import { SheetAction } from '@/types/sheet-actions';
import StepIndicator from './StepIndicator';
import ThinkingBlockView from './ThinkingBlockView';
import AnswerReveal from './AnswerReveal';
import { ResponseFooter } from './ResponseOutput';
import FollowUpsSection from './FollowUpsSection';
import QuestionChoicesPanel from './QuestionChoicesPanel';
import ActionResponseCard from './ActionResponseCard';
import GstReconCollisionCard from './GstReconCollisionCard';

/**
 * Revert affordance, rendered as one item inside the message's actions menu
 * — lets "undo this" happen right where the change was requested, instead
 * of only via the separate Change History panel or the end-of-conversation
 * LastChangeRevert bar.
 */
function TurnRevertControl({
  changeSetId,
  onRevert,
}: {
  changeSetId: string;
  onRevert: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
}) {
  const [state, setState] = useState<'idle' | 'reverting' | 'reverted' | 'error'>('idle');

  if (state === 'reverted') {
    return (
      <div className="cellix-user-msg-menu-item cellix-user-msg-menu-item-done" aria-disabled="true">
        <Check size={13} />
        <span>Reverted</span>
      </div>
    );
  }

  return (
    <button
      type="button"
      className="cellix-user-msg-menu-item"
      role="menuitem"
      disabled={state === 'reverting'}
      onClick={async () => {
        setState('reverting');
        try {
          const result = await revertChangeSet(changeSetId);
          await onRevert(changeSetId, result.inverseActions);
          setState('reverted');
        } catch (err) {
          console.error('[Cellix] Revert failed:', err);
          setState('error');
        }
      }}
    >
      <RotateCcw size={13} />
      <span>{state === 'reverting' ? 'Reverting…' : state === 'error' ? 'Retry revert' : 'Revert this change'}</span>
    </button>
  );
}

interface TurnRendererProps {
  turn: ConversationTurn;
  isActive: boolean;
  isWaiting: boolean;
  previewEnabled?: boolean;
  isApplying?: boolean;
  showActionButtons?: boolean;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onAcceptAllActions?: (turnId: string, fromBlockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onAnswerQuestion: (answer: string) => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onFollowUp: (text: string) => void;
  /** Regenerate/edit-and-resend the given turn in place (same id/position). */
  onRegenerate?: (turnId: string, overrideMessage?: string) => void;
  onRunAsAction: (message: string) => void;
  /** Powers the inline Revert control on this turn's header, when present. */
  onRevertChangeSet?: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
  /** Overwrite / Create-new button choice on a GST-recon sheet-name collision card. */
  onResolveGstReconCollision: (
    turnId: string,
    blockId: string,
    collisionId: string,
    choice: 'overwrite' | 'new',
  ) => void;
}

function PlanBlockView({
  block,
  onRunAsAction,
  disabled,
}: {
  block: PlanBlock;
  onRunAsAction: (message: string) => void;
  disabled: boolean;
}) {
  return (
    <div className="cellix-plan-card cellix-block-enter">
      <div className="cellix-plan-title">{block.type === 'plan_only' ? 'Plan (preview)' : 'Plan'}</div>
      {block.summary && <div className="cellix-plan-summary">{block.summary}</div>}
      <ol className="cellix-plan-steps">
        {block.steps.map((step, i) => (
          <li key={i}>
            <span className="cellix-plan-step-title">{step.title}</span>
            {step.detail && <span className="cellix-plan-step-detail">{step.detail}</span>}
          </li>
        ))}
      </ol>
      <div className="cellix-plan-meta">
        {block.affectedSheets.length > 0 && (
          <span className="cellix-badge">
            {block.affectedSheets.length} sheet{block.affectedSheets.length === 1 ? '' : 's'}:{' '}
            {block.affectedSheets.join(', ')}
          </span>
        )}
        {typeof block.estimatedRows === 'number' && block.estimatedRows > 0 && (
          <span className="cellix-badge">~{block.estimatedRows} rows</span>
        )}
      </div>
      {block.safestApproach && (
        <div className="cellix-plan-safest">{block.safestApproach}</div>
      )}
      {block.type === 'plan_only' && block.proposedActions && block.proposedActions.length > 0 && (
        <div className="cellix-plan-meta">
          <span className="cellix-badge">
            {block.proposedActions.length} proposed change
            {block.proposedActions.length === 1 ? '' : 's'} (not applied)
          </span>
        </div>
      )}
      <div className="cellix-action-btns">
        <button
          type="button"
          className="cellix-btn-accept"
          onClick={() => onRunAsAction(block.prompt)}
          disabled={disabled}
        >
          Run as Action
        </button>
      </div>
    </div>
  );
}

function BlockRenderer({
  block,
  turn,
  previewEnabled = false,
  isApplying = false,
  isWaiting,
  onAcceptActions,
  onAcceptAllActions,
  onRejectActions,
  onToggleThinking,
  onAnswerComplete,
  onRunAsAction,
  onAnswerQuestion,
  onResolveGstReconCollision,
  showActionButtons = true,
}: {
  block: TurnBlock;
  turn: ConversationTurn;
  previewEnabled?: boolean;
  isApplying?: boolean;
  isWaiting: boolean;
  showActionButtons?: boolean;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onAcceptAllActions?: (turnId: string, fromBlockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onRunAsAction: (message: string) => void;
  onAnswerQuestion: (answer: string) => void;
  onResolveGstReconCollision: (
    turnId: string,
    blockId: string,
    collisionId: string,
    choice: 'overwrite' | 'new',
  ) => void;
}) {
  if (block.type === 'thinking' && block.visible === false) return null;
  if (block.type === 'status' && block.visible === false) return null;

  const hideProgress =
    turn.phase === 'complete' || turn.phase === 'awaiting_input' || turn.phase === 'error';

  if (block.type === 'step') {
    if (hideProgress) {
      return null;
    }
    return <StepIndicator block={block} />;
  }

  if (block.type === 'status') {
    // Keep latest status while processing; hide after terminal phase.
    if (hideProgress) {
      return null;
    }
    return (
      <div className={`cellix-status cellix-block-enter ${block.pulsing ? 'cellix-status-pulse' : ''}`}>
        {block.pulsing ? <span className="cellix-spinner cellix-spinner-inline" /> : <div className="cellix-status-dot" />}
        <span className={block.pulsing ? 'cellix-shimmer-text' : ''}>{block.label}</span>
      </div>
    );
  }

  if (block.type === 'thinking') {
    // Always show agent thought log when it has content (Blocked / progress lines).
    if (hideProgress && !block.content?.trim()) {
      return null;
    }
    return (
      <ThinkingBlockView
        block={block}
        onToggle={() => onToggleThinking(turn.id, block.id)}
      />
    );
  }

  if (block.type === 'answer') {
    const hasActions = turn.blocks.some((b) => b.type === 'actions');
    const content = hasActions
      ? shortenActionPreviewCopy(block.content) || block.content
      : block.content;
    return (
      <AnswerReveal
        content={content}
        matches={block.matches}
        revealState={block.revealState}
        onComplete={() => onAnswerComplete(turn.id, block.id)}
        disabled={isWaiting}
        timestamp={turn.timestamp}
        showFooter={!hasActions}
      />
    );
  }

  if (block.type === 'question') {
    if (block.revealState === 'hidden') return null;
    if (turn.phase !== 'awaiting_input' || !onAnswerQuestion) return null;

    return (
      <QuestionChoicesPanel
        question={block.question}
        options={block.options}
        onSelect={onAnswerQuestion}
        disabled={isWaiting}
      />
    );
  }

  if (block.type === 'plan' || block.type === 'plan_only') {
    return (
      <PlanBlockView block={block} onRunAsAction={onRunAsAction} disabled={isWaiting} />
    );
  }

  if (block.type === 'gst_recon_collision') {
    if (!showActionButtons) return null;
    return (
      <GstReconCollisionCard
        block={block}
        disabled={isWaiting}
        onOverwrite={() =>
          onResolveGstReconCollision(turn.id, block.id, block.collisionId, 'overwrite')
        }
        onCreateNew={() =>
          onResolveGstReconCollision(turn.id, block.id, block.collisionId, 'new')
        }
      />
    );
  }

  if (block.type === 'actions') {
    const actionBlock = block as ActionBlock;
    if (actionBlock.proposalStatus === 'pending' && !showActionButtons) {
      return null;
    }
    const siblingActionBlocks = turn.blocks.filter((b) => b.type === 'actions') as ActionBlock[];
    const blockedReason = describeBlockedReason(actionBlock, siblingActionBlocks);
    const priorAnswer = turn.blocks.find((b) => b.type === 'answer');
    const priorAnswerText =
      priorAnswer && priorAnswer.type === 'answer'
        ? shortenActionPreviewCopy(priorAnswer.content) || priorAnswer.content
        : undefined;

    return (
      <ActionResponseCard
        block={actionBlock}
        previewEnabled={previewEnabled}
        isApplying={isApplying}
        showActionButtons={showActionButtons}
        blockedReason={blockedReason}
        priorAnswerText={priorAnswerText}
        onAccept={() => onAcceptActions(turn.id, block.id)}
        onAcceptAll={
          // Only offer it when there IS more after this step.
          onAcceptAllActions &&
          typeof actionBlock.stepIndex === 'number' &&
          typeof actionBlock.stepTotal === 'number' &&
          actionBlock.stepIndex < actionBlock.stepTotal
            ? () => onAcceptAllActions(turn.id, block.id)
            : undefined
        }
        onReject={() => onRejectActions(turn.id, block.id)}
      />
    );
  }

  return null;
}

/**
 * The user's own message, with a single "message actions" icon that opens a
 * dropdown (Edit and resend / Regenerate / Revert this change) rather than
 * three separate icons crowding the message. Regenerate/edit prefer
 * `onRegenerate` — it re-runs the SAME turn in place (same id, same
 * position), so the message doesn't duplicate itself further down the
 * thread the way resending as a brand new message would. Falls back to
 * `onFollowUp` (send as a new message) only if the caller didn't wire up
 * in-place regeneration.
 */
function UserMessageRow({
  turnId,
  text,
  disabled,
  onRegenerate,
  onResend,
  revertibleChangeSetIds,
  onRevertChangeSet,
}: {
  turnId: string;
  text: string;
  disabled: boolean;
  onRegenerate?: (turnId: string, overrideMessage?: string) => void;
  onResend?: (text: string) => void;
  revertibleChangeSetIds?: string[];
  onRevertChangeSet?: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
}) {
  const [isEditing, setIsEditing] = useState(false);
  const [draft, setDraft] = useState(text);
  const [menuOpen, setMenuOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const menuBoxRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!menuOpen) return;
    const handlePointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (
        triggerRef.current?.contains(target) ||
        menuBoxRef.current?.contains(target)
      ) {
        return;
      }
      setMenuOpen(false);
    };
    document.addEventListener('mousedown', handlePointerDown);
    return () => document.removeEventListener('mousedown', handlePointerDown);
  }, [menuOpen]);

  const canResend = Boolean(onRegenerate || onResend);
  const runRegenerate = (overrideMessage?: string) => {
    if (onRegenerate) {
      onRegenerate(turnId, overrideMessage);
    } else {
      onResend?.(overrideMessage ?? text);
    }
  };

  const revertControls =
    onRevertChangeSet && revertibleChangeSetIds?.length
      ? revertibleChangeSetIds.map((changeSetId) => (
          <TurnRevertControl key={changeSetId} changeSetId={changeSetId} onRevert={onRevertChangeSet} />
        ))
      : null;

  if (!canResend && !revertControls) {
    return <div className="cellix-user-msg cellix-block-enter">{text}</div>;
  }

  if (isEditing) {
    return (
      <div className="cellix-user-msg-edit cellix-block-enter">
        <textarea
          className="cellix-user-msg-edit-input"
          value={draft}
          autoFocus
          rows={Math.min(6, Math.max(2, draft.split('\n').length))}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              if (draft.trim()) {
                runRegenerate(draft.trim());
                setIsEditing(false);
              }
            } else if (e.key === 'Escape') {
              e.preventDefault();
              setDraft(text);
              setIsEditing(false);
            }
          }}
        />
        <div className="cellix-user-msg-edit-actions">
          <button
            type="button"
            className="cellix-user-msg-edit-save"
            disabled={!draft.trim()}
            onClick={() => {
              if (!draft.trim()) return;
              runRegenerate(draft.trim());
              setIsEditing(false);
            }}
          >
            <Check size={12} /> Save &amp; resend
          </button>
          <button
            type="button"
            className="cellix-user-msg-edit-cancel"
            onClick={() => {
              setDraft(text);
              setIsEditing(false);
            }}
          >
            <X size={12} /> Cancel
          </button>
        </div>
      </div>
    );
  }

  // Hidden (not just disabled) while this turn is still being processed —
  // editing/regenerating/reverting an in-flight turn isn't a real action
  // yet, so the icon only appears once there's actually something to do.
  if (disabled) {
    return <div className="cellix-user-msg cellix-block-enter">{text}</div>;
  }

  return (
    <div
      className={`cellix-user-msg-row cellix-block-enter ${
        menuOpen ? 'cellix-user-msg-row-menu-open' : ''
      }`}
    >
      <div className="cellix-user-msg cellix-user-msg-has-actions">
        <span className="cellix-user-msg-text">{text}</span>
        <div className="cellix-user-msg-menu-wrap">
          <button
            type="button"
            ref={triggerRef}
            className="cellix-user-msg-action-btn"
            aria-label="Message actions"
            title="Message actions"
            aria-haspopup="menu"
            aria-expanded={menuOpen}
            onClick={() => setMenuOpen((prev) => !prev)}
          >
            <RotateCcw size={13} />
          </button>

          {/* Floating overlay, like the reference — opens downward from the
              icon and sits on top of the content below rather than pushing
              it down. */}
          {menuOpen && (
            <div
              className="cellix-user-msg-menu"
              role="menu"
              aria-label="Message actions"
              ref={menuBoxRef}
            >
              {canResend && (
                <button
                  type="button"
                  className="cellix-user-msg-menu-item"
                  role="menuitem"
                  onClick={() => {
                    setDraft(text);
                    setIsEditing(true);
                    setMenuOpen(false);
                  }}
                >
                  <Pencil size={13} />
                  <span>Edit and resend</span>
                </button>
              )}
              {canResend && (
                <button
                  type="button"
                  className="cellix-user-msg-menu-item"
                  role="menuitem"
                  onClick={() => {
                    runRegenerate();
                    setMenuOpen(false);
                  }}
                >
                  <RefreshCw size={13} />
                  <span>Regenerate response</span>
                </button>
              )}
              {revertControls}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function blockPresentationOrder(block: TurnBlock): number {
  switch (block.type) {
    case 'step':
      return 0;
    case 'status':
      return 1;
    case 'thinking':
      return 2;
    case 'answer':
    case 'question':
      return 3;
    case 'plan':
    case 'plan_only':
      return 4;
    case 'actions':
    case 'gst_recon_collision':
      return 5;
    default:
      return 6;
  }
}

const TurnRenderer: React.FC<TurnRendererProps> = ({
  turn,
  isActive,
  isWaiting,
  previewEnabled = false,
  isApplying = false,
  showActionButtons = true,
  onAcceptActions,
  onAcceptAllActions,
  onRejectActions,
  onAnswerQuestion,
  onToggleThinking,
  onAnswerComplete,
  onFollowUp,
  onRegenerate,
  onRunAsAction,
  onRevertChangeSet,
  onResolveGstReconCollision,
}) => {
  const hideProgress = turn.phase === 'complete' || turn.phase === 'awaiting_input' || turn.phase === 'error';
  const actionDialogueReady = showActionButtons && isTurnPresentationComplete(turn);

  const revertibleActionBlocks = turn.blocks.filter(
    (b): b is ActionBlock =>
      b.type === 'actions' && b.proposalStatus === 'accepted' && Boolean(b.changeSetId),
  );

  const orderedBlocks = useMemo(
    () =>
      turn.blocks
        .map((block, index) => ({ block, index }))
        .sort((a, b) => {
          const order = blockPresentationOrder(a.block) - blockPresentationOrder(b.block);
          return order !== 0 ? order : a.index - b.index;
        })
        .map(({ block }) => block),
    [turn.blocks],
  );

  const { followUps, followUpHandler, followUpsDisabled } = useMemo(() => {
    const answerBlock = turn.blocks.find(
      (b): b is AnswerBlock => b.type === 'answer' && b.revealState === 'complete',
    );
    if (answerBlock) {
      return {
        followUps: generateSuggestedFollowUps(answerBlock.content, turn.userMessage),
        followUpHandler: onFollowUp,
        followUpsDisabled: isWaiting && isActive,
      };
    }

    return { followUps: [] as string[], followUpHandler: undefined, followUpsDisabled: true };
  }, [turn.blocks, turn.userMessage, onFollowUp, isWaiting, isActive]);

  // When a turn has an action card, its Copy + relative-time footer is
  // suppressed inside AnswerReveal/ResponseOutput (see `showFooter` above)
  // and rendered here instead, after the card rather than directly under
  // the response text — matches the reference (Claude/Cursor/Codex) layout.
  const turnHasActions = turn.blocks.some((b) => b.type === 'actions');
  const completedAnswer = turn.blocks.find(
    (b): b is AnswerBlock => b.type === 'answer' && b.revealState === 'complete',
  );

  const hasVisibleBlocks = turn.blocks.some((b) => {
    if (hideProgress && b.type === 'step') {
      return false;
    }
    if (hideProgress && b.type === 'status') {
      return false;
    }
    // Keep thinking visible after complete when the agent left a log (Blocked etc.).
    if (hideProgress && b.type === 'thinking' && !b.content?.trim()) {
      return false;
    }
    if (b.type === 'step' && b.phase === 'hidden') return false;
    if (b.type === 'answer' && b.revealState === 'hidden') return false;
    if (b.type === 'thinking' && b.visible === false) return false;
    if (b.type === 'status' && b.visible === false) return false;
    if (b.type === 'actions' && b.proposalStatus === 'pending' && !actionDialogueReady) {
      return false;
    }
    if (b.type === 'question' && turn.phase === 'awaiting_input' && b.revealState === 'visible') {
      return true;
    }
    return true;
  });

  return (
    <div className="cellix-turn">
      <UserMessageRow
        turnId={turn.id}
        text={turn.userMessage}
        disabled={isWaiting}
        onRegenerate={onRegenerate}
        onResend={onFollowUp}
        revertibleChangeSetIds={revertibleActionBlocks.map((b) => b.changeSetId!)}
        onRevertChangeSet={onRevertChangeSet}
      />

      {hasVisibleBlocks && (
        <>
          <div style={{ height: 4 }} />
          <div className="cellix-assistant-thread">
            {orderedBlocks.map((block, index) => {
              const showSep =
                index > 0 &&
                block.type === 'step' &&
                orderedBlocks[index - 1]?.type !== 'step';
              return (
                <React.Fragment key={block.id}>
                  {showSep && <div className="cellix-sep" />}
                  <BlockRenderer
                    block={block}
                    turn={turn}
                    previewEnabled={previewEnabled}
                    isApplying={isApplying}
                    isWaiting={isWaiting && isActive}
                    showActionButtons={actionDialogueReady}
                    onAcceptActions={onAcceptActions}
                    onAcceptAllActions={onAcceptAllActions}
                    onRejectActions={onRejectActions}
                    onToggleThinking={onToggleThinking}
                    onAnswerComplete={onAnswerComplete}
                    onRunAsAction={onRunAsAction}
                    onAnswerQuestion={onAnswerQuestion}
                    onResolveGstReconCollision={onResolveGstReconCollision}
                  />
                </React.Fragment>
              );
            })}
            {turnHasActions && completedAnswer && (
              <ResponseFooter content={completedAnswer.content} timestamp={turn.timestamp} />
            )}
            {followUps.length > 0 && followUpHandler && (
              <FollowUpsSection
                followUps={followUps}
                onFollowUp={followUpHandler}
                disabled={followUpsDisabled}
              />
            )}
          </div>
        </>
      )}

      {turn.error && <div className="cellix-error cellix-block-enter">{turn.error}</div>}
    </div>
  );
};

export default TurnRenderer;
