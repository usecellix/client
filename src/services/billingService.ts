import { getBillingAccountEndpoint, getBillingLedgerEndpoint } from '@/lib/apiConfig';

export interface CreditAccountSummary {
  billingEntityType: 'user' | 'org';
  planTier: 'free' | 'solo' | 'firm' | 'enterprise';
  planCredits: number;
  purchasedCredits: number;
  oneTimeCredits: number;
  availableBalance: number;
  currentPeriodEnd: string | null;
}

export interface CreditLedgerEntry {
  entryType: 'grant' | 'purchase' | 'debit' | 'one_time_grant';
  amount: number;
  bucket: 'planCredits' | 'purchasedCredits' | 'oneTimeCredits';
  actionType?: string;
  seatUserId?: string;
  createdAt: string;
}

export interface CreditLedgerPage {
  entries: CreditLedgerEntry[];
  nextCursor: string | null;
}

/** Same thin per-service fetch wrapper convention as conversationHistoryService.ts. */
async function billingFetch<T>(url: string, init?: RequestInit): Promise<T> {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    ...(init?.headers as Record<string, string> | undefined),
  };
  if (url.includes('.ngrok-free.app')) {
    headers['ngrok-skip-browser-warning'] = 'true';
  }

  const response = await fetch(url, { credentials: 'include', ...init, headers });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Billing API ${response.status}: ${text || response.statusText}`);
  }

  const body: unknown = await response.json();
  if (body && typeof body === 'object' && 'data' in body) {
    return (body as { data: T }).data;
  }
  return body as T;
}

/**
 * `GET /billing/account` provisions the Free-tier account on first read (Sept
 * 9, 2026) — it no longer 404s for a brand-new user, so the balance chip can
 * render the moment the task pane opens rather than only after a first spend.
 * The null fallback here is defensive only (an unexpected error shape), not
 * an expected steady-state outcome the way it used to be.
 */
export async function fetchCreditAccount(): Promise<CreditAccountSummary | null> {
  try {
    return await billingFetch<CreditAccountSummary>(getBillingAccountEndpoint());
  } catch (error) {
    console.warn('[Cellix] Failed to load credit account:', error);
    return null;
  }
}

export async function fetchCreditLedger(cursor?: string): Promise<CreditLedgerPage> {
  return billingFetch<CreditLedgerPage>(getBillingLedgerEndpoint(cursor));
}
