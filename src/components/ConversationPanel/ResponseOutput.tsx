import React, { useState } from 'react';
import {
  BarChart3,
  BellOff,
  Copy,
  Share2,
  ThumbsDown,
  ThumbsUp,
} from 'lucide-react';
import { MatchResult } from '@/types/conversationTurn';
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

function buildCopyText(content: string, matches?: MatchResult[]): string {
  if (!matches?.length) return content;
  const intro = stripTrailingPeriod(content.trim());
  const refs = matches.map(formatCellRef).join(', ');
  return `${intro} ${refs}.`;
}

const ResponseOutput: React.FC<ResponseOutputProps> = ({
  content,
  matches,
  followUps = [],
  onFollowUp,
  disabled = false,
  showTypingCursor = false,
  includeFollowUps = false,
}) => {
  const [copied, setCopied] = useState(false);
  const [feedback, setFeedback] = useState<'up' | 'down' | null>(null);
  const hasMatches = Boolean(matches?.length);
  const introText = hasMatches ? stripTrailingPeriod(content.trim()) : content;

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(buildCopyText(content, matches));
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1500);
    } catch {
      // ignore
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

      <div className="cellix-response-toolbar">
        <button type="button" className="cellix-toolbar-btn" onClick={handleCopy} title={copied ? 'Copied' : 'Copy'}>
          <Copy size={14} />
        </button>
        <div className="cellix-toolbar-group">
          <button
            type="button"
            className={`cellix-toolbar-btn ${feedback === 'down' ? 'active' : ''}`}
            onClick={() => setFeedback('down')}
          >
            <ThumbsDown size={14} />
          </button>
          <button
            type="button"
            className={`cellix-toolbar-btn ${feedback === 'up' ? 'active' : ''}`}
            onClick={() => setFeedback('up')}
          >
            <ThumbsUp size={14} />
          </button>
        </div>
        <button type="button" className="cellix-toolbar-btn cellix-toolbar-btn-text">
          <BarChart3 size={14} />
          <span>Usage</span>
        </button>
        <button type="button" className="cellix-toolbar-btn cellix-toolbar-btn-text">
          <Share2 size={14} />
          <span>Share</span>
        </button>
        <button type="button" className="cellix-toolbar-btn cellix-toolbar-btn-text">
          <BellOff size={14} />
          <span>Off</span>
        </button>
      </div>

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
