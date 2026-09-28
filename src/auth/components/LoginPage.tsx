import React from 'react';
import { SocialSignInButtons } from '@/auth/components/SocialSignInButtons';

const cellixLogo = new URL('../../assets/Cellix purple.png', import.meta.url).href;
const cellixWorkflowIllustration = new URL(
  '../../assets/cellix workflow without logo.png',
  import.meta.url,
).href;

/** Below this width the Google button and copy no longer fit cleanly. */
const EXPAND_BREAKPOINT_PX = 200;

interface LoginPageProps {
  error?: string | null;
  /** Called once the paired browser-tab email/password login completes (SSE push). */
  onEmailLoginComplete?: () => void;
}

export const LoginPage: React.FC<LoginPageProps> = ({ error, onEmailLoginComplete }) => {
  const [localError, setLocalError] = React.useState<string | null>(null);
  const [awaitingBrowser, setAwaitingBrowser] = React.useState(false);
  const [needsExpand, setNeedsExpand] = React.useState(false);
  const rootRef = React.useRef<HTMLDivElement>(null);
  const message = localError || error;

  React.useEffect(() => {
    const el = rootRef.current;
    if (!el || typeof ResizeObserver === 'undefined') return;

    const update = (width: number) => {
      setNeedsExpand(width > 0 && width < EXPAND_BREAKPOINT_PX);
    };

    update(el.getBoundingClientRect().width);

    const observer = new ResizeObserver((entries) => {
      const width = entries[0]?.contentRect.width ?? 0;
      update(width);
    });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  return (
    <div
      ref={rootRef}
      className={`auth-login${needsExpand ? ' auth-login--compact' : ''}`}
    >
      {needsExpand ? (
        <div className="auth-login__expand" role="status">
          <img className="auth-login__expand-logo" src={cellixLogo} alt="Cellix" />
          <div className="auth-login__expand-hint">
            <span className="auth-login__expand-arrow" aria-hidden="true">
              ←
            </span>
            <p className="auth-login__expand-text">
              Expand panel to
              <br />
              continue
            </p>
          </div>
        </div>
      ) : (
        <>
          <div className="auth-login__panel">
            <img className="auth-login__logo" src={cellixLogo} alt="Cellix" />
            <p className="auth-login__subtitle">
              Review <span>GST data</span>, match <span>ledgers</span>, validate{' '}
              <span>invoices</span>, or prepare <span>audit-ready Excel</span> work.
            </p>

            {message ? <p className="auth-login__error" role="alert">{message}</p> : null}

            {awaitingBrowser ? (
              <p className="auth-login__notice" role="status">
                Continue in the browser tab that just opened. Once you're signed in there,
                come back to Excel — this screen will switch over automatically.
              </p>
            ) : null}

            <SocialSignInButtons
              onError={(message) => {
                setLocalError(message);
                setAwaitingBrowser(false);
              }}
              onEmailLoginOpened={() => {
                setLocalError(null);
                setAwaitingBrowser(true);
              }}
              onEmailLoginComplete={onEmailLoginComplete}
            />
          </div>

          <img
            className="auth-login__illustration"
            src={cellixWorkflowIllustration}
            alt="Cellix workbook automation illustration"
          />
        </>
      )}
    </div>
  );
};
