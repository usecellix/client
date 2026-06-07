import React from 'react';
import { ChevronDown, ChevronRight } from 'lucide-react';
import { ThinkingBlock as ThinkingBlockType } from '@/types/conversationTurn';

interface ThinkingBlockViewProps {
  block: ThinkingBlockType;
  onToggle: () => void;
}

function firstLine(text: string): string {
  const line = text.split('\n').map((l) => l.trim()).find(Boolean);
  return line ?? '';
}

const ThinkingBlockView: React.FC<ThinkingBlockViewProps> = ({ block, onToggle }) => {
  const showBody = block.expanded && (block.content || block.loading);
  const preview = !block.expanded && !block.loading && block.content ? firstLine(block.content) : '';

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
      </button>

      {preview && (
        <p className="cellix-thinking-preview" title={block.content}>
          {preview}
          {block.content.includes('\n') ? '…' : ''}
        </p>
      )}

      {block.loading && !block.content && (
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
