import React from 'react';
import { Coins, X } from 'lucide-react';
import { getUpgradeUrl } from '@/lib/apiConfig';
import { openExternalUrl } from '@/utils/openExternalUrl';

interface CreditUpgradeCardProps {
  /** Credits left across all buckets. */
  balance: number;
  /**
   * Whether the balance has crossed the low-balance threshold. Drives the
   * amber warning treatment; a healthy balance stays in the brand shade so the
   * card reads as information rather than a standing alarm.
   */
  isLow?: boolean;
  onDismiss: () => void;
}


/**
 * Credit balance, seated on the composer exactly like the docked question card
 * (TASKS.md #199-#201) — same 20px inset so it is narrower than the field, same
 * rounded top with no bottom edge, so the field's own border closes it and the
 * two read as one control.
 *
 * This is the only place the balance appears while working: the top bar's chip
 * was removed in favour of it, so the number sits next to the work it is spent
 * on instead of in the chrome.
 */
const CreditUpgradeCard: React.FC<CreditUpgradeCardProps> = ({
  balance,
  isLow = false,
  onDismiss,
}) => {
  return (
    <div
      className={`cellix-credit-card cellix-block-enter${isLow ? ' is-low' : ''}`}
      role="status"
    >
      <Coins size={13} className="cellix-credit-card-icon" />
      <span className="cellix-credit-card-copy">
        <strong>{balance.toLocaleString()}</strong>{' '}
        {balance === 1 ? 'Credit' : 'Credits'} Remaining
      </span>
      <button
        type="button"
        className="cellix-credit-card-upgrade"
        onClick={() => openExternalUrl(getUpgradeUrl())}
      >
        Upgrade
      </button>
      <button
        type="button"
        className="cellix-credit-card-dismiss"
        onClick={onDismiss}
        aria-label="Dismiss credit notice"
      >
        <X size={12} />
      </button>
    </div>
  );
};

export default CreditUpgradeCard;
