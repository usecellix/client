import React, { useEffect, useState } from 'react';
import { renderLightMarkdownPlain } from '@/utils/renderLightMarkdown';
import { MatchResult } from '@/types/conversationTurn';
import ResponseOutput from './ResponseOutput';

interface AnswerRevealProps {
  content: string;
  revealState: 'hidden' | 'typing' | 'complete';
  matches?: MatchResult[];
  userPrompt?: string;
  onComplete?: () => void;
  disabled?: boolean;
}

const TYPING_INTERVAL_MS = 24;

const AnswerReveal: React.FC<AnswerRevealProps> = ({
  content,
  revealState,
  matches,
  onComplete,
  disabled = false,
}) => {
  const [displayed, setDisplayed] = useState('');
  const [isTyping, setIsTyping] = useState(false);

  const showChrome = revealState === 'complete' || (!isTyping && displayed.length > 0);

  useEffect(() => {
    if (revealState === 'hidden') {
      setDisplayed('');
      setIsTyping(false);
      return;
    }

    if (revealState === 'complete') {
      setDisplayed(content);
      setIsTyping(false);
      return;
    }

    setDisplayed('');
    setIsTyping(true);
    let index = 0;

    const timer = window.setInterval(() => {
      index += 1;
      setDisplayed(content.slice(0, index));
      if (index >= content.length) {
        window.clearInterval(timer);
        setIsTyping(false);
        onComplete?.();
      }
    }, TYPING_INTERVAL_MS);

    return () => window.clearInterval(timer);
  }, [content, revealState, onComplete]);

  if (revealState === 'hidden') return null;

  if (!showChrome) {
    return (
      <div className="cellix-response-output cellix-block-enter">
        <div className="cellix-response-shell">
          <div className={`cellix-response-text cellix-response-md ${isTyping ? 'cellix-answer-streaming' : ''}`}>
            {renderLightMarkdownPlain(displayed)}
          </div>
        </div>
      </div>
    );
  }

  return (
    <ResponseOutput
      content={displayed}
      matches={revealState === 'complete' ? matches : undefined}
      disabled={disabled}
      showTypingCursor={isTyping}
    />
  );
};

export default AnswerReveal;
