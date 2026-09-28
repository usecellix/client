import React from 'react';

const cellixLogo = new URL('../../assets/Cellix purple.png', import.meta.url).href;

interface AuthSplashProps {
  /** Override the default "Checking sign-in…" copy. */
  message?: string;
}

/** Branded loading state while the session is being verified / entering the app. */
export const AuthSplash: React.FC<AuthSplashProps> = ({ message = 'Checking sign-in…' }) => {
  return (
    <div className="auth-splash" role="status" aria-live="polite" aria-label={message}>
      <img className="auth-splash__logo" src={cellixLogo} alt="Cellix" />
      <div className="auth-splash__spinner" aria-hidden="true" />
      <p className="auth-splash__text">{message}</p>
    </div>
  );
};
