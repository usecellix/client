import { authClient } from './auth-client';

function setCopy(title: string, message: string) {
  const titleEl = document.getElementById('title');
  const messageEl = document.getElementById('message');
  if (titleEl) titleEl.textContent = title;
  if (messageEl) messageEl.textContent = message;
}

async function finishAuth() {
  const { data: session } = await authClient.getSession();
  const user = session?.user;

  const payload = JSON.stringify({
    type: 'auth-complete',
    ok: Boolean(user),
    email: user?.email ?? null,
  });

  const inOfficeDialog =
    typeof Office !== 'undefined' && Boolean(Office.context?.ui?.messageParent);

  if (user) {
    setCopy(
      'Signed in',
      inOfficeDialog
        ? 'Returning to Excel…'
        : 'You can close this browser tab and return to the Excel task pane.',
    );
  } else {
    setCopy(
      'Almost done',
      'No session cookie was found in this window. Close the tab and try signing in again from Excel.',
    );
  }

  if (inOfficeDialog) {
    try {
      Office.context.ui.messageParent(payload);
    } catch {
      // External browser — messageParent unavailable.
    }
  }
}

/* global Office */
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(() => {
    void finishAuth();
  });
} else {
  void finishAuth();
}
