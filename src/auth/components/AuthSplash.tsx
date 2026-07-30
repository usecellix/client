import React from 'react';

const cellixLogo = new URL('../../assets/Cellix purple.png', import.meta.url).href;

/** Branded loading state while the session is being verified. */
export const AuthSplash: React.FC = () => {
  return (
    <div className="auth-splash" role="status" aria-live="polite" aria-label="Checking sign-in">
      <img className="auth-splash__logo" src={cellixLogo} alt="Cellix" />
      <div className="auth-splash__spinner" aria-hidden="true" />
      <p className="auth-splash__text">Checking sign-in…</p>
    </div>
  );
};
