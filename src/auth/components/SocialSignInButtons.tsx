import React from 'react';
import { Button } from '@/components/ui/button';
import { signInWithProvider, type SocialProvider } from '@/auth/useAuth';

interface SocialSignInButtonsProps {
  disabled?: boolean;
  onError?: (message: string) => void;
}

export const SocialSignInButtons: React.FC<SocialSignInButtonsProps> = ({ disabled, onError }) => {
  const [pending, setPending] = React.useState<SocialProvider | null>(null);

  const handleSignIn = async (provider: SocialProvider) => {
    setPending(provider);
    try {
      await signInWithProvider(provider);
    } catch (error) {
      const message = error instanceof Error ? error.message : `Failed to sign in with ${provider}`;
      onError?.(message);
      setPending(null);
    }
  };

  const busy = Boolean(pending) || disabled;

  return (
    <div className="flex w-full flex-col gap-3">
      <Button
        type="button"
        variant="outline"
        className="w-full justify-center gap-2 border-[#dadce0] bg-white text-[#3c4043] hover:bg-[#f8f9fa]"
        disabled={busy}
        onClick={() => void handleSignIn('google')}
      >
        <GoogleIcon />
        {pending === 'google' ? 'Redirecting…' : 'Continue with Google'}
      </Button>

      <Button
        type="button"
        variant="outline"
        className="w-full justify-center gap-2 border-[#8c8c8c] bg-white text-[#5e5e5e] hover:bg-[#f5f5f5]"
        disabled={busy}
        onClick={() => void handleSignIn('microsoft')}
      >
        <MicrosoftIcon />
        {pending === 'microsoft' ? 'Redirecting…' : 'Continue with Microsoft'}
      </Button>
    </div>
  );
};

function GoogleIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
      <path
        fill="#4285F4"
        d="M17.64 9.2c0-.637-.057-1.251-.164-1.84H9v3.481h4.844a4.14 4.14 0 0 1-1.796 2.716v2.259h2.908c1.702-1.567 2.684-3.875 2.684-6.615z"
      />
      <path
        fill="#34A853"
        d="M9 18c2.43 0 4.467-.806 5.956-2.18l-2.908-2.259c-.806.54-1.837.86-3.048.86-2.344 0-4.328-1.584-5.036-3.711H.957v2.332A8.997 8.997 0 0 0 9 18z"
      />
      <path
        fill="#FBBC05"
        d="M3.964 10.71A5.41 5.41 0 0 1 3.682 9c0-.593.102-1.17.282-1.71V4.958H.957A8.996 8.996 0 0 0 0 9c0 1.452.348 2.827.957 4.042l3.007-2.332z"
      />
      <path
        fill="#EA4335"
        d="M9 3.58c1.321 0 2.508.454 3.44 1.345l2.582-2.58C13.463.891 11.426 0 9 0A8.997 8.997 0 0 0 .957 4.958L3.964 7.29C4.672 5.163 6.656 3.58 9 3.58z"
      />
    </svg>
  );
}

function MicrosoftIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 21 21" aria-hidden="true">
      <rect x="1" y="1" width="9" height="9" fill="#F25022" />
      <rect x="11" y="1" width="9" height="9" fill="#7FBA00" />
      <rect x="1" y="11" width="9" height="9" fill="#00A4EF" />
      <rect x="11" y="11" width="9" height="9" fill="#FFB900" />
    </svg>
  );
}
