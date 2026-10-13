import React, { useState } from 'react';
import { Check, Copy } from 'lucide-react';
import { MatchResult, formatFullDateTime, formatRelativeTime } from '@/types/conversationTurn';
import { navigateToCell } from '@/services/rangeFetchService';
import {
  renderInlineBoldMarkdown,
  renderLightMarkdownPlain,
  renderLightMarkdownWithRows,
} from '@/utils/renderLightMarkdown';

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
  /**
   * When false, the Copy + relative-time footer is omitted here so the
   * caller can render it itself, positioned after other turn content (e.g.
   * an action card) rather than directly under the response text.
   */
  showFooter?: boolean;
  /**
   * Set while an answer that names its rows is still being typed out. The
   * text shown so far may not contain a "(row N)" yet, so the layout cannot be
   * worked out from it; this says which one to use from the first character.
   */
  inlineRows?: boolean;
}

/** Copy + relative-time row — split out so it can be positioned after an
 * action card instead of always directly under the response text. */
export const ResponseFooter: React.FC<{ content: string; timestamp: Date }> = ({
  content,
  timestamp,
}) => {
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
  );
};

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

async function jumpTo(match: MatchResult): Promise<void> {
  try {
    await navigateToCell(match.sheetName, match.row, match.col, match.endCol);
  } catch (error) {
    // The sheet may have been renamed or deleted since the answer was given.
    console.warn('[Cellix] Failed to navigate to match:', error);
  }
}

/**
 * An answer that names rows as "(row 445)" and came with a pointer for each:
 * the row numbers in the text become the pointers, in place. This is the shape
 * of an answer computed over a sheet's rows ("the 10 largest debits").
 */
export function namesItsRows(content: string, matches: MatchResult[]): boolean {
  return matches.some((m) => m.rowNum !== undefined && content.includes(`(row ${m.rowNum})`));
}

function RowPointer({ match }: { match: MatchResult }) {
  return (
    <button
      type="button"
      className="cellix-row-pointer"
      onClick={() => void jumpTo(match)}
      title={`Show row ${match.rowNum} in ${match.sheetName}`}
      aria-label={`Show row ${match.rowNum} in ${match.sheetName}`}
    >
      row {match.rowNum}
    </button>
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
  showFooter = true,
  inlineRows: inlineRowsProp,
}) => {
  const inlineRows = Boolean(matches?.length) && (inlineRowsProp ?? namesItsRows(content, matches ?? []));
  // The older shape: a one-line find result followed by its list of cells.
  const hasMatches = Boolean(matches?.length) && !inlineRows;
  const introText = hasMatches ? stripTrailingPeriod(content.trim()) : content;
  const byRow = new Map((matches ?? []).filter((m) => m.rowNum !== undefined).map((m) => [m.rowNum as number, m]));

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
          ) : inlineRows ? (
            renderLightMarkdownWithRows(introText, (rowNumber, key) => {
              const match = byRow.get(rowNumber);
              return match ? <RowPointer key={key} match={match} /> : null;
            })
          ) : (
            renderLightMarkdownPlain(introText)
          )}
        </div>
      </div>

      {showFooter && timestamp && !showTypingCursor && (
        <ResponseFooter content={content} timestamp={timestamp} />
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
