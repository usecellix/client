import React, { useState } from 'react';
import { ChevronDown, ChevronRight } from 'lucide-react';

interface FollowUpsSectionProps {
  followUps: string[];
  onFollowUp?: (text: string) => void;
  disabled?: boolean;
}

const FollowUpsSection: React.FC<FollowUpsSectionProps> = ({
  followUps,
  onFollowUp,
  disabled = false,
}) => {
  const [open, setOpen] = useState(false);

  if (followUps.length === 0) return null;

  return (
    <>
      <div className="cellix-response-divider" />
      <div className="cellix-followups cellix-block-enter">
        <button
          type="button"
          className="cellix-followups-toggle"
          onClick={() => setOpen((prev) => !prev)}
          aria-expanded={open}
        >
          {open ? (
            <ChevronDown size={14} className="cellix-followups-chevron" />
          ) : (
            <ChevronRight size={14} className="cellix-followups-chevron" />
          )}
          <span>Suggested follow-ups</span>
          {!open && (
            <span className="cellix-followups-count">{followUps.length}</span>
          )}
        </button>
        {open && (
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
  );
};

export default FollowUpsSection;
