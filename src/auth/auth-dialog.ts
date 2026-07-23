import { authClient } from './auth-client';

async function startOAuth() {
  const origin = window.location.origin;
  const callbackURL = `${origin}/src/auth/auth-complete.html?provider=google`;

  const el = document.getElementById('status');
  if (el) el.textContent = 'Opening Google…';

  // Go straight to Google — account chooser (prompt=select_account
  // is configured on the server). No custom Cellix account list in this dialog.
  await authClient.signIn.social({
    provider: 'google',
    callbackURL,
  });
}

/* global Office */
Office.onReady(() => {
  void startOAuth().catch((error: unknown) => {
    const el = document.getElementById('status');
    if (el) el.textContent = error instanceof Error ? error.message : 'Sign-in failed';
  });
});
