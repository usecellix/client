import React from 'react';
import { useSession } from '@/auth/auth-client';
import { AuthSplash } from '@/auth/components/AuthSplash';
import { LoginPage } from '@/auth/components/LoginPage';

interface AuthGateProps {
  children: React.ReactNode;
}

/**
 * Gates the task pane behind a Better Auth session.
 * Splash while checking → login if signed out → app if signed in.
 */
export const AuthGate: React.FC<AuthGateProps> = ({ children }) => {
  const { data: session, isPending, error } = useSession();
  const [showSplash, setShowSplash] = React.useState(true);

  // Keep splash briefly so the UI doesn’t flash when the session resolves instantly.
  React.useEffect(() => {
    if (isPending) {
      setShowSplash(true);
      return;
    }

    const timer = window.setTimeout(() => setShowSplash(false), 400);
    return () => window.clearTimeout(timer);
  }, [isPending]);

  if (isPending || showSplash) {
    return <AuthSplash />;
  }

  if (!session) {
    return (
      <LoginPage
        error={error ? 'Unable to verify session. Please sign in again.' : null}
      />
    );
  }

  return <div className="auth-shell">{children}</div>;
};
