import React, { useState } from 'react';
import {
  BarChart3,
  BellOff,
  ChevronDown,
  ChevronRight,
  Copy,
  Share2,
  ThumbsDown,
  ThumbsUp,
} from 'lucide-react';
import { renderLightMarkdownPlain } from '@/utils/renderLightMarkdown';

interface ResponseOutputProps {
  content: string;
  followUps: string[];
  onFollowUp?: (text: string) => void;
  disabled?: boolean;
  showTypingCursor?: boolean;
}

const ResponseOutput: React.FC<ResponseOutputProps> = ({
  content,
  followUps,
  onFollowUp,
  disabled = false,
  showTypingCursor = false,
}) => {
  const [copied, setCopied] = useState(false);
  const [feedback, setFeedback] = useState<'up' | 'down' | null>(null);
  const [followUpsOpen, setFollowUpsOpen] = useState(true);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(content);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1500);
    } catch {
      // ignore
    }
  };

  return (
    <div className="cellix-response-output cellix-block-enter">
      <div className="cellix-response-shell">
        <div className={`cellix-response-text cellix-response-md ${showTypingCursor ? 'cellix-answer-streaming' : ''}`}>
          {renderLightMarkdownPlain(content)}
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

      {followUps.length > 0 && (
        <>
          <div className="cellix-response-divider" />
          <div className="cellix-followups">
            <button
              type="button"
              className="cellix-followups-toggle"
              onClick={() => setFollowUpsOpen((open) => !open)}
            >
              {followUpsOpen ? (
                <ChevronDown size={14} className="cellix-followups-chevron" />
              ) : (
                <ChevronRight size={14} className="cellix-followups-chevron" />
              )}
              <span>Suggested follow-ups</span>
            </button>
            {followUpsOpen && (
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
            )}
          </div>
        </>
      )}
    </div>
  );
};

export default ResponseOutput;
