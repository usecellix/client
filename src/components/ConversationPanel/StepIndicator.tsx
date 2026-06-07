import React from 'react';
import { Shell } from 'lucide-react';
import { StepBlock as StepBlockType } from '@/types/conversationTurn';

interface StepIndicatorProps {
  block: StepBlockType;
}

const StepIndicator: React.FC<StepIndicatorProps> = ({ block }) => {
  if (block.phase === 'hidden') return null;

  const isRunning = block.phase === 'running';
  const isDone = block.phase === 'done';
  const isRevealed = block.phase === 'revealed';
  const shimmer = isRunning || isRevealed;

  return (
    <div
      className={`cellix-step cellix-block-enter ${
        isRunning ? 'cellix-step-active' : isDone ? 'cellix-step-done' : 'cellix-step-revealed'
      }`}
    >
      <span
        className={`cellix-step-icon ${isRunning ? 'cellix-step-icon-spin' : ''}`}
        aria-hidden
      >
        <Shell size={14} strokeWidth={2.25} />
      </span>
      <span className={`cellix-step-label ${shimmer ? 'cellix-shimmer-text' : ''}`}>{block.label}</span>
      {isRevealed && (
        <span className="cellix-step-dots" aria-hidden>
          <span />
          <span />
          <span />
        </span>
      )}
      {isDone && (
        <span className="cellix-step-check" aria-hidden>
          ✓
        </span>
      )}
    </div>
  );
};

export default StepIndicator;
