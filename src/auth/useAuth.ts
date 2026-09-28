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
    return (
      typeof Office !== 'undefined' &&
      Boolean(Office.context?.ui?.displayDialogAsync)
    );
  } catch {
    return false;
  }
}

/**
 * Office Dialog = small window inside Excel → default Google OAuth page.
 * Server uses prompt=select_account so Google shows its native account chooser when
 * that WebView already has Google cookies (after the first sign-in in this dialog).
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

/**
 * Opens the marketing site's email/password login page in the system's
 * default external browser (not the Office dialog — WebView2's dialog proved
 * unreliable for messageParent/session-store updates here). Returns a random
 * token embedded in that URL — pass it to waitForEmailLogin to be told, via
 * push (SSE) rather than polling, the moment that tab finishes signing in.
 */
export function openEmailLoginPage(): string {
  const token = createLoginToken();
  const url = getEmailLoginUrl(token);
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

    const started = Date.now();
    const POLL_MS = 800;
    const TIMEOUT_MS = 2 * 60 * 1000;
    const poll = async () => {
      while (!settled && !cancelled && Date.now() - started < TIMEOUT_MS) {
        await finishWithClaim();
        if (settled || cancelled) return;
        await new Promise((r) => setTimeout(r, POLL_MS));
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

  return {
    promise,
    cancel: () => {
      if (settled) return;
      cancelled = true;
      settled = true;
      source?.close();
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
  if (!response.ok) {
    const detail = await response.text().catch(() => '');
    throw new Error(
      `Could not claim Excel login session (${response.status})${detail ? `: ${detail}` : ''}`,
    );
  }
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

export async function signOutUser(): Promise<void> {
  await authClient.signOut();
}
