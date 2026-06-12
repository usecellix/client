import React, { useState, useRef, useEffect, KeyboardEvent } from 'react';
import { ClarificationPayload } from '@/types/cellix.types';

interface ClarificationPanelProps {
  payload: ClarificationPayload;
  onAnswer: (answer: string) => void;
  onDismiss: () => void;
}

function ambiguityScoreClass(score: number): string {
  if (score > 85) return 'cellix-clarify-score-high';
  if (score > 65) return 'cellix-clarify-score-mid';
  return 'cellix-clarify-score-low';
}

export const ClarificationPanel: React.FC<ClarificationPanelProps> = ({
  payload,
  onAnswer,
  onDismiss,
}) => {
  const [customAnswer, setCustomAnswer] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  const handleCustomSubmit = () => {
    const trimmed = customAnswer.trim();
    if (trimmed) onAnswer(trimmed);
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      handleCustomSubmit();
    }
    if (e.key === 'Escape') onDismiss();
  };

  return (
    <div className="cellix-clarification-panel cellix-block-enter">
      <div className="cellix-clarify-header">
        <div className="cellix-clarify-title-row">
          <span className="cellix-clarify-icon">🤔</span>
          <span className="cellix-clarify-title">Need a bit more info</span>
        </div>
        <div className="cellix-clarify-header-actions">
          <span className={`cellix-clarify-score ${ambiguityScoreClass(payload.ambiguityScore)}`}>
            {payload.ambiguityScore}% ambiguous
          </span>
          <button
            type="button"
            onClick={onDismiss}
            className="cellix-clarify-dismiss"
            title="Dismiss"
          >
            ✕
          </button>
        </div>
      </div>

      <p className="cellix-clarify-question">{payload.question}</p>

      {payload.suggestions && payload.suggestions.length > 0 && (
        <div className="cellix-clarify-chips">
          {payload.suggestions.map((suggestion) => (
            <button
              key={suggestion}
              type="button"
              onClick={() => onAnswer(suggestion)}
              className="cellix-clarify-chip"
            >
              {suggestion}
            </button>
          ))}
        </div>
      )}

      <div className="cellix-clarify-input-row">
        <input
          ref={inputRef}
          type="text"
          value={customAnswer}
          onChange={(e) => setCustomAnswer(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder="Or type your answer..."
          className="cellix-clarify-input"
        />
        <button
          type="button"
          onClick={handleCustomSubmit}
          disabled={!customAnswer.trim()}
          className="cellix-clarify-submit"
        >
          →
        </button>
      </div>
    </div>
  );
};

export default ClarificationPanel;
