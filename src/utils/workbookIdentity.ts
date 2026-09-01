/**
 * Durable per-workbook identity (TASKS.md #22, per the #21 spike).
 *
 * Unlike `resolveWorkbookKey` (chatSessionStorage.ts — a localStorage-scoping
 * key derived from the workbook's *name*, which changes on rename), this ID
 * is minted once and persisted inside the .xlsx file itself via Office.js
 * `document.settings`, so it survives rename, close/reopen, and a 24h
 * conversation-reset — the durable correlation key ARCHITECTURE.md AD-9 and
 * DATABASE_SCHEMA.md §6.1 need for checkpoints to scope "everything that's
 * ever happened to this file."
 *
 * Per #21's finding: mint-once, get-before-set — never overwrite an existing
 * value — and treat a missing/null read as re-mintable rather than an error,
 * since a documented Excel-for-Mac bug (office-js#6091) can silently fail to
 * persist a setting despite saveAsync reporting success.
 */

const SETTING_KEY = 'cellix:workbookId';

function generateId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `wb_${crypto.randomUUID()}`;
  }
  return `wb_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

function getSettings(): Office.Settings | undefined {
  if (typeof Office === 'undefined') return undefined;
  return Office.context?.document?.settings;
}

function saveSettingsAsync(settings: Office.Settings): Promise<void> {
  return new Promise((resolve, reject) => {
    settings.saveAsync((result: Office.AsyncResult<void>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

/**
 * Resolves the durable workbookId for the currently open workbook, minting
 * and persisting one if none exists yet. Never overwrites an existing value.
 * Falls back to an unpersisted, session-local ID if `document.settings` is
 * unavailable (non-Excel host, or a saveAsync failure) rather than throwing —
 * matching `resolveWorkbookKey`'s existing fail-open convention.
 */
export async function resolveWorkbookId(): Promise<string> {
  try {
    const settings = getSettings();
    if (!settings) return generateId();

    const existing = settings.get(SETTING_KEY);
    if (typeof existing === 'string' && existing.length > 0) {
      return existing;
    }

    const minted = generateId();
    settings.set(SETTING_KEY, minted);
    try {
      await saveSettingsAsync(settings);
    } catch (error) {
      console.warn('[Cellix] Failed to persist workbookId to document settings:', error);
    }
    return minted;
  } catch (error) {
    console.warn('[Cellix] Failed to resolve workbookId, falling back to a session-local id:', error);
    return generateId();
  }
}
