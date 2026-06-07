import React from 'react';
import {
  ActionBlock,
  ConversationTurn,
  TurnBlock,
  formatMessageTime,
} from '@/types/conversationTurn';
import { SheetAction } from '@/hooks/useSseStream';
import StepIndicator from './StepIndicator';
import ThinkingBlockView from './ThinkingBlockView';
import AnswerReveal from './AnswerReveal';
import ResponseOutput from './ResponseOutput';

interface TurnRendererProps {
  turn: ConversationTurn;
  isActive: boolean;
  isWaiting: boolean;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onAnswerQuestion: (answer: string) => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onFollowUp: (text: string) => void;
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

function BlockRenderer({
  block,
  turn,
  isWaiting,
  onAcceptActions,
  onRejectActions,
  onAnswerQuestion,
  onToggleThinking,
  onAnswerComplete,
  onFollowUp,
}: {
  block: TurnBlock;
  turn: ConversationTurn;
  isWaiting: boolean;
  onAcceptActions: (turnId: string, blockId: string) => void;
  onRejectActions: (turnId: string, blockId: string) => void;
  onAnswerQuestion: (answer: string) => void;
  onToggleThinking: (turnId: string, blockId: string) => void;
  onAnswerComplete: (turnId: string, blockId: string) => void;
  onFollowUp: (text: string) => void;
}) {
  if (block.type === 'thinking' && block.visible === false) return null;
  if (block.type === 'status' && block.visible === false) return null;

  if (block.type === 'step') {
    if (turn.phase === 'complete' || turn.phase === 'awaiting_input' || turn.phase === 'error') {
      return null;
    }
    return <StepIndicator block={block} />;
  }

  if (block.type === 'status') {
    if (turn.phase === 'complete' || turn.phase === 'awaiting_input') {
      return null;
    }
  }

  if (block.type === 'thinking') {
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
        revealState={block.revealState}
        userPrompt={turn.userMessage}
        onComplete={() => onAnswerComplete(turn.id, block.id)}
        onFollowUp={onFollowUp}
        disabled={isWaiting}
      />
    );
  }

  if (block.type === 'question') {
    if (block.revealState === 'hidden') return null;

    return (
      <ResponseOutput
        content={block.question}
        followUps={block.options ?? []}
        onFollowUp={onAnswerQuestion}
        disabled={isWaiting || turn.phase !== 'awaiting_input'}
      />
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

    return (
      <div className="cellix-action-card cellix-block-enter">
        <div className="cellix-action-card-title">Preview highlighted. Cellix will make these changes:</div>
        <ul>
          <li>{actionBlock.explanation}</li>
          {actionBlock.actions.slice(0, 4).map((action, i) => (
            <li key={i}>{describeAction(action)}</li>
          ))}
          {actionBlock.actions.length > 4 && (
            <li>…and {actionBlock.actions.length - 4} more</li>
          )}
        </ul>
        {isPending && (
          <div className="cellix-action-btns">
            <button type="button" className="cellix-btn-accept" onClick={() => onAcceptActions(turn.id, block.id)}>
              Accept
            </button>
            <button type="button" className="cellix-btn-secondary" onClick={() => onAcceptActions(turn.id, block.id)}>
              Accept All
            </button>
            <button type="button" className="cellix-btn-reject" onClick={() => onRejectActions(turn.id, block.id)}>
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
  onAcceptActions,
  onRejectActions,
  onAnswerQuestion,
  onToggleThinking,
  onAnswerComplete,
  onFollowUp,
}) => {
  const hideProgress = turn.phase === 'complete' || turn.phase === 'awaiting_input' || turn.phase === 'error';

  const hasVisibleBlocks = turn.blocks.some((b) => {
    if (hideProgress && (b.type === 'step' || b.type === 'status')) return false;
    if (b.type === 'step' && b.phase === 'hidden') return false;
    if (b.type === 'answer' && b.revealState === 'hidden') return false;
    if (b.type === 'thinking' && b.visible === false) return false;
    if (b.type === 'status' && b.visible === false) return false;
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
                    isWaiting={isWaiting && isActive}
                    onAcceptActions={onAcceptActions}
                    onRejectActions={onRejectActions}
                    onAnswerQuestion={onAnswerQuestion}
                    onToggleThinking={onToggleThinking}
                    onAnswerComplete={onAnswerComplete}
                    onFollowUp={onFollowUp}
                  />
                </React.Fragment>
              );
            })}
          </div>
        </>
      )}

      {turn.error && <div className="cellix-error cellix-block-enter">{turn.error}</div>}
    </div>
  );
};

export default TurnRenderer;
