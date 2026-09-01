import React, { useEffect, useRef, useState } from 'react';
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
  const listRef = useRef<HTMLUListElement>(null);

  // This toggle is local state — it doesn't touch the conversation's turns
  // array, so the panel's sticky-to-bottom scroll never sees it expand.
  // Without this, opening it near the bottom of the scroll area reveals the
  // list below the visible viewport, hidden behind the composer.
  useEffect(() => {
    if (open) {
      listRef.current?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  }, [open]);

  if (followUps.length === 0) return null;

  return (
    <div className="cellix-followups cellix-block-enter">
      <button
        type="button"
        className="cellix-followups-toggle"
        onClick={() => setOpen((prev) => !prev)}
        aria-expanded={open}
      >
        {open ? (
          <ChevronDown size={12} className="cellix-followups-chevron" />
        ) : (
          <ChevronRight size={12} className="cellix-followups-chevron" />
        )}
        <span>Suggested follow-ups</span>
      </button>
      {open && (
        <ul className="cellix-followups-list" ref={listRef}>
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
  );
};

export default FollowUpsSection;
