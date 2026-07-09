import React, { useMemo, useRef, useState } from 'react';

const OTHER_LABEL = 'Other';

interface QuestionChoicesPanelProps {
  question: string;
  options?: string[];
  onSelect: (answer: string) => void;
  disabled?: boolean;
}

function normalizeOptions(options: string[]): string[] {
  const list = options
    .map((option) => option.trim())
    .filter((option) => option.length > 0 && option.toLowerCase() !== 'other');
  list.push(OTHER_LABEL);
  return list;
}

const QuestionChoicesPanel: React.FC<QuestionChoicesPanelProps> = ({
  question,
  options = [],
  onSelect,
  disabled = false,
}) => {
  const [showOtherInput, setShowOtherInput] = useState(options.length === 0);
  const [otherText, setOtherText] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  const choiceOptions = useMemo(() => normalizeOptions(options), [options]);

  const handleOptionClick = (option: string) => {
    if (disabled) return;
    if (option === OTHER_LABEL) {
      setShowOtherInput(true);
      window.setTimeout(() => inputRef.current?.focus(), 0);
      return;
    }
    onSelect(option);
  };

  const submitOther = () => {
    const trimmed = otherText.trim();
    if (trimmed) onSelect(trimmed);
  };

  return (
    <div className="cellix-choices cellix-block-enter">
      <p className="cellix-choices-question">{question}</p>
      {choiceOptions.length > 0 && (
        <div className="cellix-choices-list" role="listbox" aria-label="Answer choices">
          {choiceOptions.map((option) => (
            <button
              key={option}
              type="button"
              role="option"
              className={`cellix-choices-option${option === OTHER_LABEL ? ' is-other' : ''}${
                showOtherInput && option === OTHER_LABEL ? ' is-active' : ''
              }`}
              onClick={() => handleOptionClick(option)}
              disabled={disabled}
            >
              {option}
            </button>
          ))}
        </div>
      )}
      {showOtherInput && (
        <div className="cellix-choices-other-row">
          <input
            ref={inputRef}
            type="text"
            className="cellix-choices-other-input"
            value={otherText}
            onChange={(event) => setOtherText(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                submitOther();
              }
            }}
            placeholder="Type your answer…"
            disabled={disabled}
          />
          <button
            type="button"
            className="cellix-choices-other-submit"
            onClick={submitOther}
            disabled={disabled || !otherText.trim()}
            aria-label="Submit answer"
          >
            →
          </button>
        </div>
      )}
    </div>
  );
};

export default QuestionChoicesPanel;
