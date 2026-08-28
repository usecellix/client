import React, { useState } from 'react';
import { Check, Copy } from 'lucide-react';
import { MatchResult, formatFullDateTime, formatRelativeTime } from '@/types/conversationTurn';
import { navigateToCell } from '@/services/rangeFetchService';
import { renderInlineBoldMarkdown, renderLightMarkdownPlain } from '@/utils/renderLightMarkdown';

interface ResponseOutputProps {
  content: string;
  matches?: MatchResult[];
  followUps?: string[];
  onFollowUp?: (text: string) => void;
  disabled?: boolean;
  showTypingCursor?: boolean;
  /** When false, follow-ups are omitted (render at turn level via FollowUpsSection). */
  includeFollowUps?: boolean;
  /** When the response finished — powers the footer's Copy + relative-time row. */
  timestamp?: Date;
}

function formatCellRef(match: MatchResult): string {
  if (match.colLetter && match.rowNum) {
    return `${match.sheetName}!${match.colLetter}${match.rowNum}`;
  }
  return match.sheetName;
}

function FindPointersInline({ matches }: { matches: MatchResult[] }) {
  const handleNavigate = async (sheetName: string, row: number, col: number) => {
    try {
      await navigateToCell(sheetName, row, col);
    } catch (error) {
      console.warn('[Cellix] Failed to navigate to match:', error);
    }
  };

  return (
    <>
      {matches.map((match, index) => {
        const cellRef = formatCellRef(match);
        return (
          <React.Fragment key={`${match.sheetName}-${match.row}-${match.col}-${index}`}>
            {index > 0 ? <span className="cellix-find-sep">, </span> : null}
            <button
              type="button"
              className="cellix-find-pointer"
              onClick={() => handleNavigate(match.sheetName, match.row, match.col)}
              title={`Jump to ${cellRef}`}
            >
              {cellRef}
            </button>
          </React.Fragment>
        );
      })}
      <span className="cellix-find-end">.</span>
    </>
  );
}

function stripTrailingPeriod(text: string): string {
  return text.replace(/\.\s*$/, '');
}

const ResponseOutput: React.FC<ResponseOutputProps> = ({
  content,
  matches,
  followUps = [],
  onFollowUp,
  disabled = false,
  showTypingCursor = false,
  includeFollowUps = false,
  timestamp,
}) => {
  const hasMatches = Boolean(matches?.length);
  const introText = hasMatches ? stripTrailingPeriod(content.trim()) : content;
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(content);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1500);
    } catch (error) {
      console.warn('[Cellix] Copy failed:', error);
    }
  };

  return (
    <div className="cellix-response-output cellix-block-enter">
      <div className="cellix-response-shell">
        <div
          className={`cellix-response-text cellix-response-md ${hasMatches ? 'cellix-find-oneline' : ''} ${showTypingCursor ? 'cellix-answer-streaming' : ''}`}
        >
          {hasMatches && matches ? (
            <span className="cellix-find-oneline-line">
              {renderInlineBoldMarkdown(introText)}
              {' '}
              <FindPointersInline matches={matches} />
            </span>
          ) : (
            renderLightMarkdownPlain(introText)
          )}
        </div>
      </div>

      {timestamp && !showTypingCursor && (
        <div className="cellix-response-footer">
          <button
            type="button"
            className="cellix-response-footer-btn"
            aria-label={copied ? 'Copied' : 'Copy response'}
            title={copied ? 'Copied' : 'Copy response'}
            onClick={handleCopy}
          >
            {copied ? <Check size={12} /> : <Copy size={12} />}
          </button>
          <span className="cellix-response-footer-time" title={formatFullDateTime(timestamp)}>
            {formatRelativeTime(timestamp)}
          </span>
        </div>
      )}

      {includeFollowUps && followUps.length > 0 && (
        <>
          <div className="cellix-response-divider" />
          <div className="cellix-followups">
            <span className="cellix-followups-toggle" style={{ cursor: 'default' }}>
              <span>Suggested follow-ups</span>
            </span>
            <ul className="cellix-followups-list">
              {followUps.map((item, index) => (
                <li
                  key={item}
                  className={`cellix-followups-item ${index === followUps.length - 1 ? 'is-last' : ''}`}
                >
                  <button
                    type="button"
                    className="cellix-followups-link"
                    disabled={disabled}
                    onClick={() => onFollowUp?.(item)}
                  >
                    {item}
                  </button>
                </li>
              ))}
            </ul>
          </div>
        </>
      )}
    </div>
  );
};

export default ResponseOutput;
