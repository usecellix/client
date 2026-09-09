import { useCallback, useEffect, useState } from 'react';
import { CreditAccountSummary, fetchCreditAccount } from '@/services/billingService';

export interface CreditsSseEvent {
  planCredits: number;
  purchasedCredits: number;
  oneTimeCredits: number;
  debited: number;
  actionType: string;
}

interface UseCreditBalanceReturn {
  account: CreditAccountSummary | null;
  isLoading: boolean;
  /** Low-balance nudge threshold — below 20% of the plan allotment (CREDIT_SYSTEM.md §5). */
  isLowBalance: boolean;
  /**
   * Applies a `credits` SSE event (emitted once per completed debit,
   * CREDIT_SYSTEM_SCHEMA.md §6) to local state, so the indicator updates live
   * instead of only on next mount/refetch.
   */
  applyCreditsEvent: (event: CreditsSseEvent) => void;
  refresh: () => Promise<void>;
}

/**
 * Fetches the signed-in user's credit balance on mount and keeps it current
 * via `credits` SSE events streamed alongside conversation turns.
 * `GET /billing/account` provisions the Free-tier account on first read (Sept
 * 9, 2026), so a brand-new user sees their real starting balance immediately
 * rather than a blank chip until their first spend. `account: null` is now
 * purely a defensive fallback for an unexpected fetch failure, not an
 * expected steady state.
 */
export function useCreditBalance(): UseCreditBalanceReturn {
  const [account, setAccount] = useState<CreditAccountSummary | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  const refresh = useCallback(async () => {
    try {
      const summary = await fetchCreditAccount();
      setAccount(summary);
    } catch (error) {
      console.warn('[Cellix] Failed to load credit balance:', error);
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const applyCreditsEvent = useCallback((event: CreditsSseEvent) => {
    setAccount((prev) => ({
      billingEntityType: prev?.billingEntityType ?? 'user',
      planTier: prev?.planTier ?? 'free',
      planCredits: event.planCredits,
      purchasedCredits: event.purchasedCredits,
      oneTimeCredits: event.oneTimeCredits,
      availableBalance: event.planCredits + event.purchasedCredits + event.oneTimeCredits,
      currentPeriodEnd: prev?.currentPeriodEnd ?? null,
    }));
  }, []);

  const isLowBalance =
    account !== null &&
    account.planTier !== 'free' &&
    account.planCredits > 0 &&
    // 20% of the current plan allotment can't be derived from balance alone
    // without knowing the cycle's starting grant, so this approximates it
    // against the low, fixed floor that matters in practice for a nudge.
    account.availableBalance <= Math.max(20, account.planCredits * 0.2);

  return { account, isLoading, isLowBalance, applyCreditsEvent, refresh };
}
