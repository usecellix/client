// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { renderHook, waitFor } from '@testing-library/react';
import { useCreditBalance } from './useCreditBalance';
import type { CreditAccountSummary } from '@/services/billingService';

/**
 * Live bug (credit-system-v2 session): `isLowBalance` required
 * `planTier !== 'free'` AND `planCredits > 0` — both of which are ALWAYS
 * false for a Free-tier account, since its balance lives entirely in
 * `oneTimeCredits`. The low-balance nudge was silently disabled for exactly
 * the users most likely to actually run out.
 */

function accountResponse(overrides: Partial<CreditAccountSummary> = {}): CreditAccountSummary {
  return {
    billingEntityType: 'user',
    planTier: 'free',
    planCredits: 0,
    purchasedCredits: 0,
    oneTimeCredits: 120,
    availableBalance: 120,
    currentPeriodEnd: null,
    ...overrides,
  };
}

function mockFetchOnce(account: CreditAccountSummary) {
  vi.stubGlobal(
    'fetch',
    vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve(account),
    } as Response),
  );
}

describe('useCreditBalance — isLowBalance', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('is TRUE for a Free-tier account near its oneTimeCredits floor (the live bug)', async () => {
    mockFetchOnce(accountResponse({ planTier: 'free', oneTimeCredits: 15, availableBalance: 15 }));

    const { result } = renderHook(() => useCreditBalance());
    await waitFor(() => expect(result.current.isLoading).toBe(false));

    expect(result.current.isLowBalance).toBe(true);
  });

  it('is FALSE for a Free-tier account with a healthy balance', async () => {
    mockFetchOnce(accountResponse({ planTier: 'free', oneTimeCredits: 120, availableBalance: 120 }));

    const { result } = renderHook(() => useCreditBalance());
    await waitFor(() => expect(result.current.isLoading).toBe(false));

    expect(result.current.isLowBalance).toBe(false);
  });

  it('is TRUE for a Beta account near its planCredits floor', async () => {
    // The threshold approximates the cycle's starting grant using the
    // CURRENT planCredits value itself (documented pre-existing
    // approximation — the real starting grant isn't in this payload).
    // availableBalance <= max(20, planCredits * 0.2): 15 <= max(20, 3) = 20.
    mockFetchOnce(
      accountResponse({ planTier: 'beta', planCredits: 15, oneTimeCredits: 0, availableBalance: 15 }),
    );

    const { result } = renderHook(() => useCreditBalance());
    await waitFor(() => expect(result.current.isLoading).toBe(false));

    expect(result.current.isLowBalance).toBe(true);
  });

  it('is TRUE for a Solo account near its planCredits floor', async () => {
    // planCredits IS availableBalance here (no purchased/one-time credits) —
    // 15 <= max(20, 15*0.2=3) = 20.
    mockFetchOnce(
      accountResponse({ planTier: 'solo', planCredits: 15, oneTimeCredits: 0, availableBalance: 15 }),
    );

    const { result } = renderHook(() => useCreditBalance());
    await waitFor(() => expect(result.current.isLoading).toBe(false));

    expect(result.current.isLowBalance).toBe(true);
  });

  it('is FALSE for a Solo account with a healthy balance', async () => {
    // 2500 <= max(20, 2500*0.2=500) is false — comfortably above the floor.
    mockFetchOnce(
      accountResponse({ planTier: 'solo', planCredits: 2500, oneTimeCredits: 0, availableBalance: 2500 }),
    );

    const { result } = renderHook(() => useCreditBalance());
    await waitFor(() => expect(result.current.isLoading).toBe(false));

    expect(result.current.isLowBalance).toBe(false);
  });

  it('is FALSE when the account fetch fails (account stays null)', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('network error')));

    const { result } = renderHook(() => useCreditBalance());
    await waitFor(() => expect(result.current.isLoading).toBe(false));

    expect(result.current.account).toBeNull();
    expect(result.current.isLowBalance).toBe(false);
  });
});
