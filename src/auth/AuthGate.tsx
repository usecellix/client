import React from 'react';
import { useSession } from '@/auth/auth-client';
import { AuthSplash } from '@/auth/components/AuthSplash';
import { LoginPage } from '@/auth/components/LoginPage';

interface AuthGateProps {
  children: React.ReactNode;
}

/** Minimum time to show the "Signing you in…" splash after email login claim. */
const ENTER_APP_MIN_MS = 1000;

function waitAtLeast(startedAt: number, minMs: number): Promise<void> {
  const remaining = minMs - (Date.now() - startedAt);
  if (remaining <= 0) return Promise.resolve();
  return new Promise((resolve) => window.setTimeout(resolve, remaining));
}

/**
 * Gates the task pane behind a Better Auth session.
 * Splash while checking → login if signed out → app if signed in.
 */
export const AuthGate: React.FC<AuthGateProps> = ({ children }) => {
  const { data: session, isPending, error, refetch } = useSession();
  const [showSplash, setShowSplash] = React.useState(true);
  const [enteringApp, setEnteringApp] = React.useState(false);
  const [claimError, setClaimError] = React.useState<string | null>(null);
  // Only the very first resolution should show the splash/loading state.
  // Better Auth's useSession() re-flips isPending to true on every refetch()
  // call while signed out (data is null), so without this flag the
  // background poll below would flicker the splash screen every tick.
  const hasResolvedOnce = React.useRef(false);
  if (!isPending) hasResolvedOnce.current = true;

  const sessionRef = React.useRef(session);
  sessionRef.current = session;

  // The claim error only describes a failed sign-in attempt. Clear it whenever
  // a user is present, and again on sign-out (user -> none) so it can never
  // show on the login screen after logging out.
  const hadUserRef = React.useRef(false);
  React.useEffect(() => {
    const hasUser = Boolean(session?.user);
    if (hasUser || hadUserRef.current) setClaimError(null);
    hadUserRef.current = hasUser;
  }, [session]);

  // Keep splash briefly so the UI doesn’t flash when the session resolves instantly.
  React.useEffect(() => {
    if (isPending && !hasResolvedOnce.current) {
      setShowSplash(true);
      return;
    }

    const timer = window.setTimeout(() => setShowSplash(false), 400);
    return () => window.clearTimeout(timer);
  }, [isPending]);

  // The session is checked once on load (useSession). Email/password sign-in
  // happens in a browser tab, but waitForEmailLogin (useAuth.ts) claims the
  // session and calls handleEmailLoginComplete itself, so no background
  // polling is needed here. Only re-check when the user returns to Excel —
  // event-driven, not on a timer.
  React.useEffect(() => {
    if (session?.user || enteringApp) return;

    const onReturn = () => {
      if (document.visibilityState === 'visible') void refetch();
    };
    window.addEventListener('focus', onReturn);
    document.addEventListener('visibilitychange', onReturn);

    return () => {
      window.removeEventListener('focus', onReturn);
      document.removeEventListener('visibilitychange', onReturn);
    };
  }, [session, refetch, enteringApp]);

  const handleEmailLoginComplete = React.useCallback(async () => {
    setEnteringApp(true);
    const startedAt = Date.now();
    setClaimError(null);
    try {
      // Claim already set the cookie — refetch until the session store sees it.
      let signedIn = false;
      for (let i = 0; i < 6; i++) {
        // refetch is typed `void` but resolves the session at runtime.
        const result: unknown = await refetch();
        const user =
          result && typeof result === 'object' && 'data' in result
            ? (result as { data?: { user?: unknown } | null }).data?.user
            : undefined;
        if (user) {
          signedIn = true;
          break;
        }
        await new Promise((r) => window.setTimeout(r, 200));
      }
      await waitAtLeast(startedAt, ENTER_APP_MIN_MS);
      // Decide only after the wait, against the LIVE session (the refetch
      // result above doesn't always expose the user even when it landed).
      // Error only if the pane truly has no session — the browser refused to
      // store the cookie (third-party/blocked cookies).
      if (!signedIn && !sessionRef.current?.user) {
        setClaimError(
          'You signed in, but Excel could not keep the session. Allow cookies for this add-in (or turn off "block third-party cookies" for Excel) and try again.',
        );
      }
    } finally {
      setEnteringApp(false);
    }
  }, [refetch]);

  if (enteringApp) {
    return <AuthSplash message="Signing you in…" />;
  }

  if (showSplash && !hasResolvedOnce.current) {
    return <AuthSplash />;
  }

  if (!session) {
    return (
      <LoginPage
        error={
          claimError ??
          (error
            ? "Can't reach the server right now. Check your connection and try again."
            : null)
        }
        onEmailLoginComplete={() => void handleEmailLoginComplete()}
      />
    );
  }

  return <div className="auth-shell">{children}</div>;
};
