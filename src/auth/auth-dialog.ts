import { authClient } from './auth-client';

type SocialProvider = 'google' | 'microsoft';

function setStatus(message: string) {
  const el = document.getElementById('status');
  if (el) el.textContent = message;
}

async function startOAuth() {
  const params = new URLSearchParams(window.location.search);
  const provider: SocialProvider =
    params.get('provider') === 'microsoft' ? 'microsoft' : 'google';

  const origin = window.location.origin;
  const callbackURL = `${origin}/src/auth/auth-complete.html?provider=${encodeURIComponent(provider)}`;

  setStatus(provider === 'microsoft' ? 'Opening Microsoft…' : 'Opening Google…');

  // Go straight to the provider — Google’s default account chooser (prompt=select_account
  // is configured on the server). No custom Cellix account list in this dialog.
  await authClient.signIn.social({
    provider,
    callbackURL,
  });
}

/* global Office */
Office.onReady(() => {
  void startOAuth().catch((error: unknown) => {
    setStatus(error instanceof Error ? error.message : 'Sign-in failed');
  });
});
