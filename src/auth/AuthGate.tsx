import React from 'react';
import { useSession } from '@/auth/auth-client';
import { LoginPage } from '@/auth/components/LoginPage';

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

  return <div className="auth-shell">{children}</div>;
};
