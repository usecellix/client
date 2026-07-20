import React from 'react';
import { useSession } from '@/auth/auth-client';
import { LoginPage } from '@/auth/components/LoginPage';
import { signOutUser } from '@/auth/useAuth';
import { Button } from '@/components/ui/button';

interface AuthGateProps {
  children: React.ReactNode;
}

/**
 * Gates the task pane behind a Better Auth session.
 * Shows Google / Microsoft login until the user is authenticated.
 */
export const AuthGate: React.FC<AuthGateProps> = ({ children }) => {
  const { data: session, isPending, error } = useSession();

  if (isPending) {
    return (
      <div className="auth-login">
        <p className="auth-login__subtitle">Checking session…</p>
      </div>
    );
  }

  if (!session) {
    return (
      <LoginPage
        error={error ? 'Unable to verify session. Please sign in again.' : null}
      />
    );
  }

  return (
    <div className="auth-shell">
      <header className="auth-shell__bar">
        <div className="auth-shell__user">
          <span className="auth-shell__name">{session.user.name || 'Signed in'}</span>
          {session.user.email ? (
            <span className="auth-shell__email">{session.user.email}</span>
          ) : null}
        </div>
        <Button type="button" variant="ghost" size="sm" onClick={() => void signOutUser()}>
          Sign out
        </Button>
      </header>
      <div className="auth-shell__content">{children}</div>
    </div>
  );
};
