export { authClient, useSession, signIn, signOut } from './auth-client';
export { AuthGate } from './AuthGate';
export { AuthSplash } from './components/AuthSplash';
export { LoginPage } from './components/LoginPage';
export { SocialSignInButtons } from './components/SocialSignInButtons';
export {
  signInWithProvider,
  signOutUser,
  getAuthCompleteUrl,
  getAuthDialogUrl,
} from './useAuth';
export type { SocialProvider } from './useAuth';
