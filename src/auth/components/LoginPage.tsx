import React from 'react';
import { SocialSignInButtons } from '@/auth/components/SocialSignInButtons';

interface LoginPageProps {
  error?: string | null;
}

export const LoginPage: React.FC<LoginPageProps> = ({ error }) => {
  const [localError, setLocalError] = React.useState<string | null>(null);
  const message = localError || error;

  return (
    <div className="auth-login">
      <div className="auth-login__panel">
        <p className="auth-login__brand">Cellix</p>
        <h1 className="auth-login__title">Sign in</h1>
        <p className="auth-login__subtitle">
          Use your Google or Microsoft account to continue in the Excel task pane.
        </p>

        {message ? <p className="auth-login__error" role="alert">{message}</p> : null}

        <SocialSignInButtons onError={setLocalError} />
      </div>
    </div>
  );
};
