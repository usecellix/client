import React, { useEffect, useMemo, useState } from 'react';
import { generateSuggestedFollowUps } from '@/utils/suggestedFollowUps';
import { renderLightMarkdownPlain } from '@/utils/renderLightMarkdown';
import ResponseOutput from './ResponseOutput';

interface AnswerRevealProps {
  content: string;
  revealState: 'hidden' | 'typing' | 'complete';
  userPrompt?: string;
  onComplete?: () => void;
  onFollowUp?: (text: string) => void;
  disabled?: boolean;
}

const TYPING_INTERVAL_MS = 24;

const AnswerReveal: React.FC<AnswerRevealProps> = ({
  content,
  revealState,
  userPrompt,
  onComplete,
  onFollowUp,
  disabled = false,
}) => {
  const [displayed, setDisplayed] = useState('');
  const [isTyping, setIsTyping] = useState(false);

  const followUps = useMemo(
    () => generateSuggestedFollowUps(content, userPrompt),
    [content, userPrompt],
  );

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
      followUps={followUps}
      onFollowUp={onFollowUp}
      disabled={disabled}
      showTypingCursor={isTyping}
    />
  );
};

export default AnswerReveal;
