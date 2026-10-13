/* global Office */

/**
 * Opens a URL in the user's real browser.
 *
 * A task pane is an embedded webview: `window.open` there can silently do
 * nothing, or worse navigate the pane itself and take the add-in down with it.
 * `Office.context.ui.openBrowserWindow` is the supported way to hand a URL to
 * the system browser, with `window.open` kept only as a fallback for when the
 * Office host predates it. TASKS.md #201.
 */
export function openExternalUrl(url: string): void {
  try {
    const officeUi = (globalThis as { Office?: typeof Office }).Office?.context?.ui;
    if (officeUi?.openBrowserWindow) {
      officeUi.openBrowserWindow(url);
      return;
    }
  } catch (error) {
    console.warn('[Cellix] openBrowserWindow failed, falling back to window.open:', error);
  }
  window.open(url, '_blank', 'noopener,noreferrer');
}
