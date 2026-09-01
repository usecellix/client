import { authClient } from './auth-client';

export type SocialProvider = 'google';

export function getAuthCompleteUrl(provider?: SocialProvider): string {
  const origin =
    typeof window !== 'undefined' ? window.location.origin : 'https://localhost:3000';
  const base = `${origin}/src/auth/auth-complete.html`;
  return provider ? `${base}?provider=${encodeURIComponent(provider)}` : base;
}

export function getAuthDialogUrl(provider: SocialProvider): string {
  const origin =
    typeof window !== 'undefined' ? window.location.origin : 'https://localhost:3000';
  return `${origin}/src/auth/auth-dialog.html?provider=${encodeURIComponent(provider)}`;
}

function canUseOfficeDialog(): boolean {
  try {
    return (
      typeof Office !== 'undefined' &&
      Boolean(Office.context?.ui?.displayDialogAsync)
    );
  } catch {
    return false;
  }
}

/**
 * Office Dialog = small window inside Excel → default Google OAuth page.
 * Server uses prompt=select_account so Google shows its native account chooser when
 * that WebView already has Google cookies (after the first sign-in in this dialog).
 */
export async function signInWithProvider(provider: SocialProvider): Promise<void> {
  if (canUseOfficeDialog()) {
    await signInWithOfficeDialog(provider);
    return;
  }

  await authClient.signIn.social({
    provider,
    callbackURL: getAuthCompleteUrl(provider),
  });
}

function signInWithOfficeDialog(provider: SocialProvider): Promise<void> {
  const dialogUrl = getAuthDialogUrl(provider);

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 70, width: 45, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(asyncResult.error?.message || 'Could not open sign-in dialog'));
          return;
        }

        const dialog = asyncResult.value;

        // Office types both dialog events with one handler signature taking the
        // union of payloads, so each handler must accept it and narrow.
        type DialogArg = { message: string; origin: string | undefined } | { error: number };

        const onMessage = (arg: DialogArg) => {
          if (!('message' in arg)) return;
          dialog.close();
          try {
            const data = JSON.parse(arg.message) as { type?: string; ok?: boolean };
            if (data.type === 'auth-complete' && data.ok) {
              window.location.reload();
              resolve();
              return;
            }
            reject(new Error('Sign-in did not complete. Please try again.'));
          } catch {
            reject(new Error('Invalid sign-in response from dialog'));
          }
        };

        const onEvent = (arg: DialogArg) => {
          if (!('error' in arg)) return;
          if (arg.error === 12006) {
            reject(new Error('Sign-in was cancelled'));
            return;
          }
          reject(new Error(`Sign-in dialog error (${arg.error})`));
        };

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, onMessage);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, onEvent);
      },
    );
  });
}

export async function signOutUser(): Promise<void> {
  await authClient.signOut();
}
