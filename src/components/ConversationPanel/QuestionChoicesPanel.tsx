import React, { useEffect, useMemo, useRef, useState } from 'react';
import { ChevronDown, ChevronUp, MessageCircleQuestion } from 'lucide-react';

const OTHER_LABEL = 'Other';

/** Letter badges shown against each choice, Cursor-style: A, B, C… */
const OPTION_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

/** Sent when the user skips rather than choosing — the turn is still waiting on a reply. */
export const SKIP_ANSWER = 'Skip the question and proceed with your best judgment.';

interface QuestionChoicesPanelProps {
  question: string;
  options?: string[];
  onSelect: (answer: string) => void;
  disabled?: boolean;
  /**
   * Rendered docked at the bottom of the panel rather than inline in the
   * transcript. The docked variant is the full Cursor-style card — header,
   * lettered choices, keyboard shortcuts, Skip/Continue — and confirms a
   * choice rather than submitting it on the first click, so a selection can be
   * combined with typed detail. Inline keeps the original submit-on-click
   * behaviour, since answered questions in the transcript are a record, not a
   * control. TASKS.md #182, #186.
   */
  docked?: boolean;
  /** Docked only — lifted so the composer can submit selection + typed detail together. */
  selected?: string | null;
  onSelectedChange?: (option: string | null) => void;
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
  docked = false,
  selected = null,
  onSelectedChange,
}) => {
  const [showOtherInput, setShowOtherInput] = useState(options.length === 0);
  const [otherText, setOtherText] = useState('');
  const [collapsed, setCollapsed] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const choiceOptions = useMemo(() => normalizeOptions(options), [options]);

  // "Other" is always appended, so a card with nothing but it is a free-text
  // question with no choices to make.
  const freeTextOnly = choiceOptions.length === 1;

  // Docked derives the input from the live selection rather than latching a
  // flag: picking "Other" and then changing to a real choice used to leave the
  // input stranded on screen, adding a box the question never asked for.
  const showDockedInput = freeTextOnly || selected === OTHER_LABEL;

  // A free-text question has nothing to click, so focus has to land in the
  // input or the card arrives inert.
  useEffect(() => {
    if (docked && showDockedInput) inputRef.current?.focus();
  }, [docked, showDockedInput]);

  const selectOption = (option: string) => {
    if (disabled) return;
    onSelectedChange?.(option);
    if (option === OTHER_LABEL) {
      window.setTimeout(() => inputRef.current?.focus(), 0);
    }
  };

  const handleOptionClick = (option: string) => {
    if (disabled) return;
    // Inline: the click IS the answer, as before.
    if (!docked) {
      if (option === OTHER_LABEL) {
        setShowOtherInput(true);
        window.setTimeout(() => inputRef.current?.focus(), 0);
        return;
      }
      onSelect(option);
      return;
    }
    selectOption(option);
  };

  const resolveAnswer = (): string | null => {
    const typed = otherText.trim();
    if (selected && selected !== OTHER_LABEL) {
      return typed ? `${selected} — ${typed}` : selected;
    }
    return typed || null;
  };

  const submit = () => {
    const answer = resolveAnswer();
    if (answer) onSelect(answer);
  };

  const submitOther = () => {
    const trimmed = otherText.trim();
    if (!trimmed) return;
    onSelect(docked ? (resolveAnswer() ?? trimmed) : trimmed);
  };

  const canContinue = Boolean(resolveAnswer());

  // Keyboard, Cursor-style: letters pick a choice, Enter confirms, Esc skips.
  // Letter and Enter handling is suppressed whenever focus sits in a text field
  // — the composer below this card is live, and typing "a" there must type an
  // "a", not silently choose option A.
  useEffect(() => {
    if (!docked || disabled || collapsed) return;

    const isTypingTarget = (target: EventTarget | null): boolean => {
      const el = target as HTMLElement | null;
      if (!el) return false;
      const tag = el.tagName;
      return tag === 'INPUT' || tag === 'TEXTAREA' || el.isContentEditable;
    };

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.metaKey || event.ctrlKey || event.altKey) return;

      // Every shortcut, Escape included, defers to a focused text field. The
      // composer below owns Escape for dismissing its @-reference menu, and
      // skipping the question out from under that would be a nasty surprise.
      // The Skip button stays clickable either way.
      if (isTypingTarget(event.target)) return;

      if (event.key === 'Escape') {
        event.preventDefault();
        onSelect(SKIP_ANSWER);
        return;
      }

      if (event.key === 'Enter') {
        if (!canContinue) return;
        event.preventDefault();
        submit();
        return;
      }

      if (event.key.length === 1) {
        const index = OPTION_LETTERS.indexOf(event.key.toUpperCase());
        if (index >= 0 && index < choiceOptions.length) {
          event.preventDefault();
          selectOption(choiceOptions[index]);
        }
      }
    };

    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  });

  if (!docked) {
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
                aria-selected={showOtherInput && option === OTHER_LABEL}
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
  }

  return (
    <div className="cellix-choices cellix-block-enter is-docked">
      <div className="cellix-choices-head">
        <div className="cellix-choices-head-title">
          <MessageCircleQuestion size={13} />
          <span>Questions</span>
        </div>
        <button
          type="button"
          className="cellix-choices-collapse"
          onClick={() => setCollapsed((value) => !value)}
          aria-expanded={!collapsed}
          aria-label={collapsed ? 'Expand question' : 'Collapse question'}
        >
          {collapsed ? <ChevronDown size={13} /> : <ChevronUp size={13} />}
          <span>1 of 1</span>
        </button>
      </div>

      {!collapsed && (
        <>
          <div className="cellix-choices-prompt">
            <span className="cellix-choices-index">1.</span>
            <p className="cellix-choices-question">{question}</p>
          </div>

          <div className="cellix-choices-list" role="listbox" aria-label="Answer choices">
            {choiceOptions.map((option, index) => (
              <button
                key={option}
                type="button"
                role="option"
                aria-selected={selected === option}
                className={`cellix-choices-option${selected === option ? ' is-selected' : ''}`}
                onClick={() => handleOptionClick(option)}
                disabled={disabled}
              >
                <span className="cellix-choices-key" aria-hidden="true">
                  {OPTION_LETTERS[index] ?? '•'}
                </span>
                <span className="cellix-choices-option-text">{option}</span>
              </button>
            ))}
          </div>

          {showDockedInput && (
            <input
              ref={inputRef}
              type="text"
              className="cellix-choices-other-input is-docked"
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
          )}

          <div className="cellix-choices-actions">
            <button
              type="button"
              className="cellix-choices-skip"
              onClick={() => onSelect(SKIP_ANSWER)}
              disabled={disabled}
            >
              Skip <kbd>Esc</kbd>
            </button>
            <button
              type="button"
              className="cellix-choices-continue"
              onClick={submit}
              disabled={disabled || !canContinue}
            >
              Continue <kbd>⏎</kbd>
            </button>
          </div>
        </>
      )}
    </div>
  );
};

export default QuestionChoicesPanel;
