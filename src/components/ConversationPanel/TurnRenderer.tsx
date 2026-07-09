import React, { useMemo } from 'react';
import {
  ActionBlock,
  AnswerBlock,
  ConversationTurn,
  PlanBlock,
  TurnBlock,
  formatMessageTime,
} from '@/types/conversationTurn';
import { SheetAction } from '@/hooks/useSseStream';
import { generateSuggestedFollowUps } from '@/utils/suggestedFollowUps';
import StepIndicator from './StepIndicator';
import ThinkingBlockView from './ThinkingBlockView';
import AnswerReveal from './AnswerReveal';
import FollowUpsSection from './FollowUpsSection';
import QuestionChoicesPanel from './QuestionChoicesPanel';

interface TurnRendererProps {
  turn: ConversationTurn;
  isActive: boolean;
  isWaiting: boolean;
  previewEnabled?: boolean;
  isApplying?: boolean;
  showActionButtons?: boolean;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onAnswerQuestion: (answer: string) => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onFollowUp: (text: string) => void;
  onRunAsAction: (message: string) => void;
}

function describeAction(action: SheetAction): string {
  if (action.type === 'SET_CELL' && action.row !== undefined && action.col !== undefined) {
    const col = String.fromCharCode(65 + action.col);
    return `Set cell ${col}${action.row + 1} to ${action.value ?? ''}`;
  }
  if (action.type === 'CLEAR_CELL' && action.row !== undefined && action.col !== undefined) {
    const col = String.fromCharCode(65 + action.col);
    return `Clear cell ${col}${action.row + 1}`;
  }
  if (action.type === 'SET_FORMULA' && action.row !== undefined && action.col !== undefined) {
    const col = String.fromCharCode(65 + action.col);
    return `Set formula in ${col}${action.row + 1}`;
  }
  if (action.type === 'HIGHLIGHT_CELL' && action.row !== undefined && action.col !== undefined) {
    const col = String.fromCharCode(65 + action.col);
    return `Highlight cell ${col}${action.row + 1}`;
  }
  if (action.type === 'ADD_ROW') return 'Add a new row';
  if (action.type === 'DELETE_ROW' && action.row !== undefined) return `Delete row ${action.row + 1}`;
  return action.type;
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
      <div className="cellix-plan-title">Plan</div>
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
  onRejectActions,
  onToggleThinking,
  onAnswerComplete,
  onRunAsAction,
  onAnswerQuestion,
  showActionButtons = true,
}: {
  block: TurnBlock;
  turn: ConversationTurn;
  previewEnabled?: boolean;
  isApplying?: boolean;
  isWaiting: boolean;
  showActionButtons?: boolean;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onRunAsAction: (message: string) => void;
  onAnswerQuestion: (answer: string) => void;
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
    if (hideProgress) {
      return null;
    }
  }

  if (block.type === 'thinking') {
    if (hideProgress) {
      return null;
    }
    return (
      <ThinkingBlockView
        block={block}
        onToggle={() => onToggleThinking(turn.id, block.id)}
      />
    );
  }

  if (block.type === 'status') {
    return (
      <div className={`cellix-status cellix-block-enter ${block.pulsing ? 'cellix-status-pulse' : ''}`}>
        {block.pulsing ? <span className="cellix-spinner cellix-spinner-inline" /> : <div className="cellix-status-dot" />}
        <span className={block.pulsing ? 'cellix-shimmer-text' : ''}>{block.label}</span>
      </div>
    );
  }

  if (block.type === 'answer') {
    return (
      <AnswerReveal
        content={block.content}
        matches={block.matches}
        revealState={block.revealState}
        onComplete={() => onAnswerComplete(turn.id, block.id)}
        disabled={isWaiting}
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

  if (block.type === 'plan') {
    return (
      <PlanBlockView block={block} onRunAsAction={onRunAsAction} disabled={isWaiting} />
    );
  }

  if (block.type === 'actions') {
    const actionBlock = block as ActionBlock;
    const isPending = actionBlock.proposalStatus === 'pending';
    const isAccepted = actionBlock.proposalStatus === 'accepted';

    if (isAccepted) {
      return (
        <div className="cellix-changes-card cellix-block-enter">
          <div style={{ fontSize: 12.5, fontWeight: 500, color: 'var(--cx-gray-700)' }}>
            {actionBlock.explanation}
          </div>
          <div className="cellix-changes-link">
            <span className="cellix-badge">
              {actionBlock.actions.length} Direct Change
              {actionBlock.actions.length === 1 ? '' : 's'}
            </span>
            <span>Applied</span>
          </div>
        </div>
      );
    }

    if (actionBlock.proposalStatus === 'rejected') {
      return (
        <div className="cellix-changes-card cellix-block-enter" style={{ opacity: 0.7 }}>
          <div className="cellix-action-card-title">Changes rejected</div>
          <div style={{ fontSize: 12, color: 'var(--cx-gray-500)' }}>{actionBlock.explanation}</div>
        </div>
      );
    }

    if (isPending && previewEnabled) {
      return (
        <div className="cellix-changes-card cellix-block-enter">
          <div className="cellix-changes-summary">{actionBlock.explanation}</div>
          <div className="cellix-changes-link">
            <span className="cellix-badge">
              {actionBlock.actions.length} Direct Change
              {actionBlock.actions.length === 1 ? '' : 's'}
            </span>
            <span>Review below</span>
          </div>
        </div>
      );
    }

    return (
      <div className="cellix-action-card cellix-block-enter">
        <div className="cellix-action-card-title">Cellix will make these changes:</div>
        <ul>
          <li>{actionBlock.explanation}</li>
          {actionBlock.actions.slice(0, 4).map((action, i) => (
            <li key={i}>{describeAction(action)}</li>
          ))}
          {actionBlock.actions.length > 4 && (
            <li>…and {actionBlock.actions.length - 4} more</li>
          )}
        </ul>
        {isPending && showActionButtons && (
          <div className="cellix-action-btns">
            <button
              type="button"
              className="cellix-btn-accept"
              onClick={() => onAcceptActions(turn.id, block.id)}
              disabled={isApplying}
            >
              {isApplying ? 'Applying…' : 'Accept'}
            </button>
            <button
              type="button"
              className="cellix-btn-reject"
              onClick={() => onRejectActions(turn.id, block.id)}
              disabled={isApplying}
            >
              Reject
            </button>
          </div>
        )}
      </div>
    );
  }

  return null;
}

const TurnRenderer: React.FC<TurnRendererProps> = ({
  turn,
  isActive,
  isWaiting,
  previewEnabled = false,
  isApplying = false,
  showActionButtons = true,
  onAcceptActions,
  onRejectActions,
  onAnswerQuestion,
  onToggleThinking,
  onAnswerComplete,
  onFollowUp,
  onRunAsAction,
}) => {
  const hideProgress = turn.phase === 'complete' || turn.phase === 'awaiting_input' || turn.phase === 'error';

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

  const hasVisibleBlocks = turn.blocks.some((b) => {
    if (hideProgress && (b.type === 'step' || b.type === 'status' || b.type === 'thinking')) {
      return false;
    }
    if (b.type === 'step' && b.phase === 'hidden') return false;
    if (b.type === 'answer' && b.revealState === 'hidden') return false;
    if (b.type === 'thinking' && b.visible === false) return false;
    if (b.type === 'status' && b.visible === false) return false;
    if (b.type === 'question' && turn.phase === 'awaiting_input' && b.revealState === 'visible') {
      return true;
    }
    return true;
  });

  return (
    <div className="cellix-turn">
      <div className="cellix-msg-meta">
        Action &nbsp;|&nbsp; {formatMessageTime(turn.timestamp)}
      </div>
      <div className="cellix-user-msg cellix-block-enter">{turn.userMessage}</div>

      {hasVisibleBlocks && (
        <>
          <div style={{ height: 4 }} />
          <div className="cellix-assistant-thread">
            {turn.blocks.map((block, index) => {
              const showSep =
                index > 0 &&
                block.type === 'step' &&
                turn.blocks[index - 1]?.type !== 'step';
              return (
                <React.Fragment key={block.id}>
                  {showSep && <div className="cellix-sep" />}
                  <BlockRenderer
                    block={block}
                    turn={turn}
                    previewEnabled={previewEnabled}
                    isApplying={isApplying}
                    isWaiting={isWaiting && isActive}
                    showActionButtons={showActionButtons}
                    onAcceptActions={onAcceptActions}
                    onRejectActions={onRejectActions}
                    onToggleThinking={onToggleThinking}
                    onAnswerComplete={onAnswerComplete}
                    onRunAsAction={onRunAsAction}
                    onAnswerQuestion={onAnswerQuestion}
                  />
                </React.Fragment>
              );
            })}
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
