import { authClient } from './auth-client';
import {
  getEmailLoginUrl,
  getExcelLoginClaimUrl,
  getExcelLoginWaitUrl,
} from '@/lib/apiConfig';

export type SocialProvider = 'google';

export function getAuthCompleteUrl(provider?: SocialProvider): string {
  const origin =
    typeof window !== 'undefined' ? window.location.origin : 'https://localhost:3000';
  const base = `${origin}/src/auth/auth-complete.html`;
  return provider ? `${base}?provider=${encodeURIComponent(provider)}` : base;
}

export function getAuthDialogUrl(provider: SocialProvider): string {
  const origin =
    typeof window !== 'undefined' ? window.location.origin : 'https://localhost:3000';
  return `${origin}/src/auth/auth-dialog.html?provider=${encodeURIComponent(provider)}`;
}

function canUseOfficeDialog(): boolean {
  try {
    return typeof Office !== 'undefined' && Boolean(Office.context?.ui?.displayDialogAsync);
  } catch {
    return false;
  }
}

/**
 * REINTRODUCED AT USER REQUEST (see git log bbd1705 for the removal commit
 * and its reasoning) — Google sign-in via Office's displayDialogAsync.
 *
 * KNOWN BROKEN for real users: the dialog is served from the add-in's own
 * origin, but the session cookie Better Auth's OAuth callback sets lands on
 * BETTER_AUTH_URL's origin (the marketing site). auth-complete.ts's
 * authClient.getSession() call inside the dialog can never see that
 * cross-origin cookie, so it always reports no user, the dialog times out,
 * and Office surfaces error 12006 ("Sign-in was cancelled") even though the
 * Google login itself succeeded. This is exactly what AppSource
 * certification (Product ID b5d15e34-0e2e-4680-8791-6316c46b28d7, reviewed
 * Oct 7 2026) failed on. Restored anyway per explicit instruction — do not
 * re-remove this without being asked; if sign-in reliability regresses,
 * route back through openEmailLoginPage(provider) + waitForEmailLogin +
 * POST /excel-login/claim instead, which sidesteps the cross-origin cookie
 * problem entirely (still used below for "Log in with email and password").
 */
export async function signInWithProvider(provider: SocialProvider): Promise<void> {
  if (canUseOfficeDialog()) {
    await signInWithOfficeDialog(provider);
    return;
  }

  await authClient.signIn.social({
    provider,
    callbackURL: getAuthCompleteUrl(provider),
  });
}

function signInWithOfficeDialog(provider: SocialProvider): Promise<void> {
  return openAuthDialog(getAuthDialogUrl(provider), { height: 70, width: 45 });
}

function openAuthDialog(
  dialogUrl: string,
  size: { height: number; width: number },
): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { ...size, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(asyncResult.error?.message || 'Could not open sign-in dialog'));
          return;
        }

        const dialog = asyncResult.value;

        // Office types both dialog events with one handler signature taking the
        // union of payloads, so each handler must accept it and narrow.
        type DialogArg = { message: string; origin: string | undefined } | { error: number };

        const onMessage = (arg: DialogArg) => {
          if (!('message' in arg)) return;
          dialog.close();
          try {
            const data = JSON.parse(arg.message) as { type?: string; ok?: boolean };
            if (data.type === 'auth-complete' && data.ok) {
              window.location.reload();
              resolve();
              return;
            }
            reject(new Error('Sign-in did not complete. Please try again.'));
          } catch {
            reject(new Error('Invalid sign-in response from dialog'));
          }
        };

        const onEvent = (arg: DialogArg) => {
          if (!('error' in arg)) return;
          if (arg.error === 12006) {
            reject(new Error('Sign-in was cancelled'));
            return;
          }
          reject(new Error(`Sign-in dialog error (${arg.error})`));
        };

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, onMessage);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, onEvent);
      },
    );
  });
}

/**
 * Opens the marketing site's login page in the system's default external
 * browser — still used for "Log in with email and password" (see
 * SocialSignInButtons.tsx). Google no longer calls this; see
 * signInWithProvider above. Returns a random token embedded in that URL —
 * pass it to waitForEmailLogin to be told, via push (SSE) rather than
 * polling, the moment that tab finishes signing in. `provider`, when given,
 * makes the page start that OAuth provider immediately instead of showing
 * the sign-in picker.
 */
export function openEmailLoginPage(provider?: SocialProvider): string {
  const token = createLoginToken();
  const url = getEmailLoginUrl(token, provider);
  try {
    if (typeof Office !== 'undefined' && Office.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(url);
      return token;
    }
  } catch {
    // Fall through to a plain window.open below.
  }
  window.open(url, '_blank', 'noopener,noreferrer');
  return token;
}

function createLoginToken(): string {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

/**
 * Waits for the browser tab opened by openEmailLoginPage to finish sign-in,
 * then claims a session cookie into this WebView.
 *
 * Excel WebView2 + Vite often close EventSource with `onerror` when the Nest
 * SSE stream completes — that used to reject with "Login wait connection closed"
 * *before* claim ran, even though Landing already called /excel-login/complete.
 * So we treat SSE as a hint and **poll claim** until it succeeds (or timeout).
 */
export function waitForEmailLogin(token: string): { promise: Promise<void>; cancel: () => void } {
  let cancelled = false;
  let settled = false;
  let source: EventSource | null = null;
  let resolveCleanup: () => void = () => {};
  const cleanupSignal = new Promise<void>((r) => {
    resolveCleanup = r;
  });

  const promise = new Promise<void>((resolve, reject) => {
    const finishWithClaim = async () => {
      if (settled || cancelled) return;
      try {
        await claimExcelLoginSession(token);
      } catch {
        return; // claim not ready yet — keep polling / waiting for SSE
      }
      if (settled || cancelled) return;
      settled = true;
      source?.close();
      resolve();
    };

    try {
      source = new EventSource(getExcelLoginWaitUrl(token));
    } catch {
      source = null;
    }

    if (source) {
      source.addEventListener('login-complete', () => {
        void finishWithClaim();
      });
      // Some hosts deliver Nest SSE as the default `message` event.
      source.addEventListener('message', () => {
        void finishWithClaim();
      });
      source.onerror = () => {
        // Normal when the server completes the stream after notify — try claim.
        void finishWithClaim();
      };
    }

    // Long enough for a new user to register *and then* log in on the browser
    // tab — 2 minutes was routinely exceeded by that path, so the pane gave up
    // before the tab ever reported success. Keep in sync with the Server's
    // TOKEN_TTL_MS (excel-login-bridge.service.ts).
    const TIMEOUT_MS = 15 * 60 * 1000;
    const FAST_POLL_MS = 1000;
    const SLOW_POLL_MS = 3000;
    const FAST_PHASE_MS = 20 * 1000;
    const started = Date.now();

    // Returning to Excel is the moment the user expects the switch — claim then
    // instead of waiting for the next poll tick.
    const onReturn = () => {
      if (document.visibilityState === 'visible') void finishWithClaim();
    };
    window.addEventListener('focus', onReturn);
    document.addEventListener('visibilitychange', onReturn);
    const detachReturn = () => {
      window.removeEventListener('focus', onReturn);
      document.removeEventListener('visibilitychange', onReturn);
    };
    void cleanupSignal.then(detachReturn);

    const poll = async () => {
      while (!settled && !cancelled && Date.now() - started < TIMEOUT_MS) {
        await finishWithClaim();
        if (settled || cancelled) return;
        const elapsed = Date.now() - started;
        await new Promise((r) => setTimeout(r, elapsed < FAST_PHASE_MS ? FAST_POLL_MS : SLOW_POLL_MS));
      }
      if (!settled && !cancelled) {
        settled = true;
        source?.close();
        reject(
          new Error(
            'Login wait timed out — finish signing in on the browser tab, then try again.',
          ),
        );
      }
    };
    void poll();
  });

  void promise.then(resolveCleanup, resolveCleanup);

  return {
    promise,
    cancel: () => {
      if (settled) return;
      cancelled = true;
      settled = true;
      source?.close();
      resolveCleanup();
    },
  };
}

/**
 * Redeems the pairing token for a Better Auth Set-Cookie in this WebView.
 * Must run in the Excel task pane (same origin as /api), not the Landing tab.
 */
export async function claimExcelLoginSession(token: string): Promise<void> {
  const response = await fetch(getExcelLoginClaimUrl(), {
    method: 'POST',
    credentials: 'include',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ token }),
  });
  // 202 = the browser tab hasn't finished signing in yet (normal while waiting).
  if (response.status === 202 || !response.ok) {
    const detail = await response.text().catch(() => '');
    throw new Error(
      `Could not claim Excel login session (${response.status})${detail ? `: ${detail}` : ''}`,
    );
  }
}

export async function signOutUser(): Promise<void> {
  await authClient.signOut();
}
