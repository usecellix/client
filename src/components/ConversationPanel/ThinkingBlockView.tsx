import React from 'react';
import { ChevronDown, ChevronRight } from 'lucide-react';
import { ThinkingBlock as ThinkingBlockType } from '@/types/conversationTurn';

interface ThinkingBlockViewProps {
  block: ThinkingBlockType;
  onToggle: () => void;
}

function paragraphs(text: string): string[] {
  return text.split('\n\n').map((p) => p.trim()).filter(Boolean);
}

const ThinkingBlockView: React.FC<ThinkingBlockViewProps> = ({ block, onToggle }) => {
  // Strictly tap-to-expand: collapsed shows only the toggle row, nothing
  // else — no preview snippet, no skeleton mockup. Content only appears once
  // the user has explicitly opened it via `onToggle`.
  const showBody = block.expanded && (block.content || block.loading);
  const stepCount = block.content ? paragraphs(block.content).length : 0;

  return (
    <div className="cellix-block-enter cellix-thinking-block">
      <button type="button" className="cellix-thinking-toggle" onClick={onToggle}>
        {block.loading ? (
          <span className="cellix-spinner" />
        ) : block.expanded ? (
          <ChevronDown size={12} color="#9CA3AF" />
        ) : (
          <ChevronRight size={12} color="#9CA3AF" />
        )}
        <span className={block.loading ? 'cellix-shimmer-text' : ''}>
          {block.loading ? 'Thinking…' : 'Thought process'}
        </span>
        {block.loading && stepCount > 1 && (
          <span className="cellix-thinking-step-count">Step {stepCount}</span>
        )}
      </button>

      {block.expanded && block.loading && !block.content && (
        <div className="cellix-thinking-skeleton">
          <div className="cellix-skeleton-line" style={{ width: '92%' }} />
          <div className="cellix-skeleton-line" style={{ width: '78%' }} />
          <div className="cellix-skeleton-line" style={{ width: '65%' }} />
        </div>
      )}

      {showBody && block.content && (
        <div
          className={`cellix-thinking-body ${block.loading ? 'cellix-thinking-body-loading' : ''}`}
        >
          {block.content.split('\n\n').map((paragraph, index) => (
            <p
              key={index}
              className={block.loading && index === 0 ? 'cellix-shimmer-text' : undefined}
            >
              {paragraph}
            </p>
          ))}
        </div>
      )}
    </div>
  );
};

export default ThinkingBlockView;
