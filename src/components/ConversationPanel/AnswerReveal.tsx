import React, { useEffect, useRef, useState } from 'react';
import { MatchResult } from '@/types/conversationTurn';
import ResponseOutput, { namesItsRows } from './ResponseOutput';

interface AnswerRevealProps {
  content: string;
  revealState: 'hidden' | 'typing' | 'complete';
  matches?: MatchResult[];
  userPrompt?: string;
  onComplete?: () => void;
  disabled?: boolean;
  /** When the response finished — powers the footer's Copy + relative-time row. */
  timestamp?: Date;
  /** When false, the caller renders the Copy + relative-time footer itself, positioned elsewhere. */
  showFooter?: boolean;
}

/**
 * THE typing speed — one setting for every answer in the app (reconciliations, GST, imports, ordinary
 * replies). Each tick reveals TYPING_CHARS_PER_TICK characters every TYPING_INTERVAL_MS, i.e. about
 * 200 characters a second, whatever the length of the reply. Change the speed here and nowhere else;
 * no feature needs its own setting.
 */
export const TYPING_INTERVAL_MS = 24;
export const TYPING_CHARS_PER_TICK = 5;

/**
 * A "(row 445)" that is only partly typed: "(", "(row", "(row 44". Held back
 * until it is whole, so the pointer appears in one piece instead of as
 * characters that then turn into a button.
 */
const PARTIAL_ROW_MARKER = /\((?:r(?:o(?:w(?: \d*)?)?)?)?$/;

const AnswerReveal: React.FC<AnswerRevealProps> = ({
  content,
  revealState,
  matches,
  onComplete,
  disabled = false,
  timestamp,
  showFooter = true,
}) => {
  const [displayed, setDisplayed] = useState('');
  const [isTyping, setIsTyping] = useState(false);
  // Callers pass a new function on every render. Held in a ref so that a
  // re-render of the thread while an answer is being typed (a credits update,
  // the end-of-stream event) does not restart the typing from the first
  // character.
  const onCompleteRef = useRef(onComplete);
  onCompleteRef.current = onComplete;

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
      index += TYPING_CHARS_PER_TICK;
      setDisplayed(content.slice(0, index));
      if (index >= content.length) {
        window.clearInterval(timer);
        setIsTyping(false);
        onCompleteRef.current?.();
      }
    }, TYPING_INTERVAL_MS);

    return () => window.clearInterval(timer);
  }, [content, revealState]);

  if (revealState === 'hidden') return null;

  // Decided from the whole answer, not from what has been typed so far.
  const rowsMode = Boolean(matches?.length) && namesItsRows(content, matches ?? []);
  const shown = rowsMode && isTyping ? displayed.replace(PARTIAL_ROW_MARKER, '') : displayed;

  // One renderer from the first character to the last. Typing used to be drawn
  // by a simpler one and swapped for this at the end, which replayed the
  // entrance animation and made the finished answer appear to be written twice.
  return (
    <ResponseOutput
      content={shown}
      // Row pointers appear as each line is typed. A find result's list of
      // cells is still added once the sentence before it is complete.
      matches={rowsMode || revealState === 'complete' ? matches : undefined}
      inlineRows={rowsMode ? true : undefined}
      disabled={disabled}
      showTypingCursor={isTyping}
      timestamp={timestamp}
      showFooter={showFooter}
    />
  );
};

export default AnswerReveal;
