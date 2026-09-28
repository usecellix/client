import { useState } from 'react';
import type { RepairRequest } from '@/services/repairRequest';
import {
  clearSpillBlockers as defaultClearSpillBlockers,
  describeCells,
  type SpillBlockage,
} from '@/services/spillBlockers';

/**
 * The red card under a turn — TASKS.md #321.
 *
 * It used to be bare text: "Applied, but 1 formula cell(s) returned an error:
 * Main!A19 #SPILL!", with nothing to do about it. The read-back already knew
 * more — a repair prompt (#168) was built and never rendered, and a blocked
 * spill has a precise, deterministic fix. This card offers whichever applies.
 * Both are explicit clicks: the write already landed, so changing the user's
 * cells again is their call.
 */
interface TurnErrorCardProps {
  message: string;
  spillBlockages?: SpillBlockage[];
  repairSuggestion?: RepairRequest;
  onRepair?: (prompt: string) => void;
  disabled?: boolean;
  /** Injected in tests; defaults to the real Office.js clear. */
  clearSpillBlockers?: (blockage: SpillBlockage) => Promise<boolean>;
}

type ClearState =
  | { status: 'idle' }
  | { status: 'clearing' }
  | { status: 'done' }
  | { status: 'still-blocked' }
  | { status: 'failed'; reason: string };

export function TurnErrorCard({
  message,
  spillBlockages,
  repairSuggestion,
  onRepair,
  disabled = false,
  clearSpillBlockers = defaultClearSpillBlockers,
}: TurnErrorCardProps) {
  const [clear, setClear] = useState<ClearState>({ status: 'idle' });
  const blockages = spillBlockages ?? [];
  const allBlockers = blockages.flatMap((b) => b.blockers);

  if (clear.status === 'done') {
    const anchors = blockages.map((b) => `${b.sheet}!${b.anchor}`).join(', ');
    return (
      <div className="cellix-error-card cellix-error-card--resolved cellix-block-enter" role="status">
        Cleared {describeCells(allBlockers)} — the list at {anchors} now fills in.
      </div>
    );
  }

  const handleClear = async () => {
    setClear({ status: 'clearing' });
    try {
      let allResolved = true;
      for (const blockage of blockages) {
        const resolved = await clearSpillBlockers(blockage);
        allResolved = allResolved && resolved;
      }
      setClear(allResolved ? { status: 'done' } : { status: 'still-blocked' });
    } catch (error) {
      setClear({
        status: 'failed',
        reason: error instanceof Error ? error.message : 'Excel could not clear those cells',
      });
    }
  };

  return (
    <div className="cellix-error cellix-error-card cellix-block-enter" role="alert">
      <div className="cellix-error-card-message">{message}</div>

      {blockages.length > 0 && (
        <div className="cellix-error-card-actions">
          <button
            type="button"
            className="cellix-btn-secondary"
            disabled={disabled || clear.status === 'clearing'}
            onClick={() => void handleClear()}
          >
            {clear.status === 'clearing' ? 'Clearing…' : `Clear ${describeCells(allBlockers)}`}
          </button>
          <span className="cellix-error-card-hint">Removes only those cells' contents.</span>
        </div>
      )}
      {clear.status === 'still-blocked' && (
        <div className="cellix-error-card-hint">
          Cleared, but the list still can't fill in — something else below it is in the way.
        </div>
      )}
      {clear.status === 'failed' && (
        <div className="cellix-error-card-hint">Couldn't clear those cells: {clear.reason}</div>
      )}

      {blockages.length === 0 && repairSuggestion && onRepair && (
        <div className="cellix-error-card-actions">
          <button
            type="button"
            className="cellix-btn-secondary"
            disabled={disabled}
            onClick={() => onRepair(repairSuggestion.prompt)}
          >
            {repairSuggestion.cellCount === 1 ? 'Fix this formula' : 'Fix these formulas'}
          </button>
        </div>
      )}
    </div>
  );
}

export default TurnErrorCard;
