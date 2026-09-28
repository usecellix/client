import React from 'react';
import { PlayCircle, X } from 'lucide-react';

interface ResumeRunCardProps {
  /** The wave the build would carry on FROM, already 1-based. */
  waveIndex: number;
  waveTotal: number;
  /** Continue the existing run. Disabled while something else is in flight. */
  onResume: () => void;
  onDismiss: () => void;
  disabled?: boolean;
}

/**
 * Offers to carry on a build whose connection was lost —
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 8, client half of TASKS.md #293.
 *
 * Observed live: the task pane lost its connection mid-build, the run sat in
 * `awaiting_decision` forever, and nothing on either side could resume it. A
 * stepwise build spends its whole life handing control back to this panel
 * between waves, so any interruption strands it with the finished waves
 * already applied and no route back — the only option was to rebuild the
 * whole thing, which for a twelve-month workbook is several minutes and a
 * fresh set of chances to fail.
 *
 * Deliberately a statement with two choices rather than a modal: the user may
 * well have closed the pane on purpose, and a build they abandoned should not
 * block the panel until they acknowledge it. Dismissing leaves the run
 * untouched on the server rather than cancelling it.
 *
 * Seated on the composer exactly like `CreditUpgradeCard`, so the two read as
 * one system rather than two unrelated notices.
 */
const ResumeRunCard: React.FC<ResumeRunCardProps> = ({
  waveIndex,
  waveTotal,
  onResume,
  onDismiss,
  disabled = false,
}) => {
  // A run whose total is unknown still deserves the offer — it just cannot
  // say where in the build it stopped.
  const position =
    waveTotal > 0 && waveIndex > 0 ? `step ${waveIndex} of ${waveTotal}` : 'where it stopped';

  return (
    <div className="cellix-credit-card cellix-block-enter" role="status">
      <PlayCircle size={13} className="cellix-credit-card-icon" />
      <span className="cellix-credit-card-copy">
        A build stopped at <strong>{position}</strong>
      </span>
      <button
        type="button"
        className="cellix-credit-card-upgrade"
        onClick={onResume}
        disabled={disabled}
      >
        Continue
      </button>
      <button
        type="button"
        className="cellix-credit-card-dismiss"
        onClick={onDismiss}
        aria-label="Dismiss unfinished build"
      >
        <X size={12} />
      </button>
    </div>
  );
};

export default ResumeRunCard;
