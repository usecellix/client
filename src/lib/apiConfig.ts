const DEFAULT_API_BASE_URL = '/api';
const STREAM_PATH = '/excel-ai/process';
const CONVERSATION_PATH = '/excel-ai/conversation';
const CONVERSATIONS_PATH = '/excel-ai/conversations';
const COMPARE_PATH = '/sheets/compare';

function normalizeBaseUrl(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) {
    throw new Error('VITE_API_BASE_URL is set but empty. Provide a valid URL like /api');
  }
  return trimmed.replace(/\/+$/, '');
}

export function getApiBaseUrl(): string {
  const envValue = (import.meta as any)?.env?.VITE_API_BASE_URL as string | undefined;

  if (typeof envValue === 'string') {
    return normalizeBaseUrl(envValue);
  }

  if ((import.meta as any)?.env?.DEV) {
    // eslint-disable-next-line no-console
    console.warn(
      `[Cellix] VITE_API_BASE_URL is not set. Falling back to ${DEFAULT_API_BASE_URL}. ` +
        `Create frontend/.env with VITE_API_BASE_URL=/api to use the dev proxy.`,
    );
  }

  return DEFAULT_API_BASE_URL;
}

/** Marketing site. Swap for the pricing page once it exists. */
const DEFAULT_UPGRADE_URL = 'https://www.usecellix.com';

/**
 * Where the Upgrade button sends people.
 *
 * Defaults in CODE rather than only in `.env`, because Vite freezes
 * `import.meta.env` when the dev server starts and never re-reads `.env` on
 * HMR — an env-only value meant the button stayed invisible after being
 * configured until someone restarted the server, which is a confusing way for
 * a button to be missing. `VITE_UPGRADE_URL` still overrides per environment.
 * TASKS.md #204.
 */
export function getUpgradeUrl(): string {
  const envValue = (import.meta as any)?.env?.VITE_UPGRADE_URL as string | undefined;
  const trimmed = typeof envValue === 'string' ? envValue.trim() : '';
  return trimmed || DEFAULT_UPGRADE_URL;
}

export function getStreamEndpoint(): string {
  return `${getApiBaseUrl()}${STREAM_PATH}`;
}

export function getConversationEndpoint(): string {
  return `${getApiBaseUrl()}${CONVERSATION_PATH}`;
}

export function getConversationByIdEndpoint(conversationId: string): string {
  return `${getApiBaseUrl()}${CONVERSATION_PATH}/${encodeURIComponent(conversationId)}`;
}

/** Server-backed chat history for the signed-in user (TASKS.md #171/#172). */
export function getConversationListEndpoint(params?: {
  limit?: number;
  cursor?: string;
  workbookId?: string;
}): string {
  const query = new URLSearchParams();
  if (params?.limit !== undefined) query.set('limit', String(params.limit));
  if (params?.cursor) query.set('cursor', params.cursor);
  if (params?.workbookId) query.set('workbookId', params.workbookId);
  const suffix = query.toString();
  return `${getApiBaseUrl()}${CONVERSATIONS_PATH}${suffix ? `?${suffix}` : ''}`;
}

export function getToolResultEndpoint(): string {
  return `${getApiBaseUrl()}${CONVERSATION_PATH}/tool-result`;
}

/** Advances a step-wise Tier 3 run (TASKS.md #153, STEPWISE_EXECUTION.md §3). */
export function getContinueRunEndpoint(): string {
  return `${getApiBaseUrl()}${CONVERSATION_PATH}/continue`;
}

export function getCompareEndpoint(): string {
  return `${getApiBaseUrl()}${COMPARE_PATH}`;
}

export function getGstReconcileEndpoint(): string {
  return `${getApiBaseUrl()}/gst/reconcile`;
}

/** Bank statement import: decoded file in, sheet actions out. ATTACHMENT_EXTRACTION_PLAN.md §6. */
export function getBankStatementImportEndpoint(): string {
  return `${getApiBaseUrl()}/ingest/bank-statement`;
}

export function getAuditApplyEndpoint(changeSetId: string): string {
  return `${getApiBaseUrl()}/audit/apply/${encodeURIComponent(changeSetId)}`;
}

export function getAuditPreviewLocalEndpoint(): string {
  return `${getApiBaseUrl()}/audit/preview-local`;
}

export function getAuditRevertEndpoint(changeSetId: string): string {
  return `${getApiBaseUrl()}/audit/revert/${encodeURIComponent(changeSetId)}`;
}

export function getAuditHistoryEndpoint(conversationId: string): string {
  return `${getApiBaseUrl()}/audit/history/${encodeURIComponent(conversationId)}`;
}

export function getAuditStatsEndpoint(from?: string, to?: string): string {
  const params = new URLSearchParams();
  if (from) params.set('from', from);
  if (to) params.set('to', to);
  const query = params.toString();
  return `${getApiBaseUrl()}/audit/stats${query ? `?${query}` : ''}`;
}

export function getAuditExportEndpoint(format: 'json' | 'csv', from?: string, to?: string): string {
  const params = new URLSearchParams({ format });
  if (from) params.set('from', from);
  if (to) params.set('to', to);
  return `${getApiBaseUrl()}/audit/export?${params.toString()}`;
}

export function getCheckpointCreateEndpoint(): string {
  return `${getApiBaseUrl()}/audit/checkpoint`;
}

export function getCheckpointListEndpoint(workbookId: string): string {
  return `${getApiBaseUrl()}/audit/checkpoint/${encodeURIComponent(workbookId)}`;
}

export function getCheckpointRestoreEndpoint(checkpointId: string): string {
  return `${getApiBaseUrl()}/audit/checkpoint/restore/${encodeURIComponent(checkpointId)}`;
}

/** Current credit balance (all 3 buckets) + plan tier — CREDIT_SYSTEM_SCHEMA.md §5. */
export function getBillingAccountEndpoint(): string {
  return `${getApiBaseUrl()}/billing/account`;
}

/** Paginated credit ledger history — CREDIT_SYSTEM_SCHEMA.md §5. */
export function getBillingLedgerEndpoint(cursor?: string): string {
  const query = new URLSearchParams();
  if (cursor) query.set('cursor', cursor);
  const suffix = query.toString();
  return `${getApiBaseUrl()}/billing/ledger${suffix ? `?${suffix}` : ''}`;
}

/**
 * Buys a one-time top-up pack — authed (uses the task pane's existing
 * session), unlike subscribing/upgrading. A top-up debits/credits a
 * SPECIFIC signed-in user's account, so — unlike "Upgrade the plan," which
 * can redirect to the marketing site's guest-checkout flow — this can't be a
 * plain external redirect: the marketing site has no session with this
 * backend to authenticate the call. Returns a Razorpay short_url for the
 * caller to open in a new tab for the actual payment.
 */
export function getBillingTopupEndpoint(): string {
  return `${getApiBaseUrl()}/billing/checkout/topup`;
}

// Office's displayDialogAsync refuses non-HTTPS dialog URLs, so this must be
// HTTPS even in dev — Landing-page/vite.config.ts serves HTTPS locally using
// the same office-addin-dev-certs cert this add-in's own dev server uses.
const DEFAULT_MARKETING_SITE_URL = 'https://localhost:5173';

/**
 * The marketing site (CELLIX-landing-page) — where the task pane's
 * "Upgrade the plan" action AND the email/password sign-in flow
 * (getEmailLoginUrl) send the user, since both are marketing-site-driven
 * (pricing/checkout, and /login for the paired-tab sign-in handshake).
 * Matches the backend's own CHECKOUT_SUCCESS_URL/CHECKOUT_CANCEL_URL default
 * dev port.
 *
 * A missing VITE_MARKETING_SITE_URL in a PRODUCTION build must fail loudly,
 * not fall back to localhost: that fallback silently sent a real AppSource
 * reviewer's browser to their own machine when this var was never set in
 * Vercel's project settings (the .env.production file that documents the
 * var is gitignored and local-only — committing it would not have helped;
 * Vercel still needs the var set in its own dashboard). A thrown error here
 * fails the build/first render instead of shipping a silent dead end.
 */
export function getMarketingSiteUrl(): string {
  const envValue = (import.meta as any)?.env?.VITE_MARKETING_SITE_URL as string | undefined;
  if (typeof envValue === 'string' && envValue.trim()) {
    return envValue.trim().replace(/\/+$/, '');
  }
  if (!(import.meta as any)?.env?.DEV) {
    throw new Error(
      '[Cellix] VITE_MARKETING_SITE_URL is not set in this production build. Set it in the ' +
        'Vercel project\'s Environment Variables (Production scope) to https://www.usecellix.com ' +
        '— without it, sign-in and upgrade links point at localhost for every real user.',
    );
  }
  return DEFAULT_MARKETING_SITE_URL;
}

/** Opens the marketing site's pricing page — the entry point for upgrading/changing plans. */
export function getPricingPageUrl(): string {
  return `${getMarketingSiteUrl()}/pricing`;
}

/**
 * Marketing site's login page, opened in the system's default browser for
 * EVERY sign-in method (email/password, Google) — see
 * frontend/src/auth/useAuth.ts openEmailLoginPage/signInWithProvider.
 * `from=excel` tells the page to show a "go back to Excel" message instead
 * of redirecting to /app on success; `token` is a random id the page
 * reports back to the Server on success (POST /excel-login/complete),
 * which pushes the result to the Excel add-in's matching SSE connection
 * (GET /excel-login/wait) — see AuthGate.tsx. `provider`, when set, makes
 * the login page auto-trigger that OAuth provider on load instead of
 * showing the picker (CELLIX-landing-page/src/pages/LoginPage.tsx).
 *
 * Google specifically no longer goes through Office's displayDialogAsync +
 * a same-origin-cookie read: BETTER_AUTH_URL is a single fixed origin
 * (www.usecellix.com, so the marketing site's own OAuth login works), which
 * means the OAuth callback's session cookie is set on THAT origin — a
 * dialog page served from app.usecellix.com can never see it, so
 * Office.context.ui.messageParent never fired and every Google sign-in
 * reported "Sign-in was cancelled" (Office error 12006) even after a
 * successful Google login. Confirmed in AppSource certification feedback
 * (Product ID b5d15e34-0e2e-4680-8791-6316c46b28d7, Oct 7 2026). Routing
 * through this same external-browser pairing-token flow — already proven
 * for email/password — sidesteps the cross-origin cookie problem entirely:
 * nothing needs to read a cookie across origins, the token hand-off is
 * origin-independent by design.
 */
export function getEmailLoginUrl(loginToken: string, provider?: 'google'): string {
  const base = `${getMarketingSiteUrl()}/login?from=excel&token=${encodeURIComponent(loginToken)}`;
  return provider ? `${base}&provider=${encodeURIComponent(provider)}` : base;
}

/**
 * SSE endpoint the task pane waits on for the paired browser-tab login to
 * complete. Deliberately under /excel-login, not /auth — client/vite.config.ts's
 * dev proxy treats /api/auth/* as a single kept-prefix rule pointed at Better
 * Auth's own /api/auth/* mount, so a Nest route at /auth/* would collide
 * with it through the proxy. /excel-login/* instead goes through the
 * general /api/* rule (prefix stripped), landing on the Nest route directly.
 */
export function getExcelLoginWaitUrl(loginToken: string): string {
  return `${getApiBaseUrl()}/excel-login/wait?token=${encodeURIComponent(loginToken)}`;
}

/**
 * After SSE login-complete, the task pane redeems the pairing token for a
 * Better Auth session cookie in the Office WebView jar (which never received
 * the Landing-tab Set-Cookie).
 */
export function getExcelLoginClaimUrl(): string {
  return `${getApiBaseUrl()}/excel-login/claim`;
}

/** Opens the marketing site's signed-in dashboard — the entry point for account/settings. */
export function getAccountSettingsUrl(): string {
  return `${getMarketingSiteUrl()}/app`;
}

