import { createAuthClient } from 'better-auth/react';

/**
 * Better Auth client — talks to Nest via Vite `/api/auth` proxy (same origin cookies).
 */
export const authClient = createAuthClient({
  baseURL: typeof window !== 'undefined' ? window.location.origin : 'https://localhost:3000',
});

export const { signIn, signOut, useSession } = authClient;
