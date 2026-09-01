import { afterEach, describe, expect, it, vi } from 'vitest';
import { resolveWorkbookId } from '@/utils/workbookIdentity';

function makeMockSettings(initial: Record<string, unknown> = {}) {
  const store = new Map<string, unknown>(Object.entries(initial));
  const saveAsync = vi.fn((callback: (result: { status: string; error?: unknown }) => void) => {
    callback({ status: 'succeeded' });
  });
  return {
    get: vi.fn((name: string) => store.get(name)),
    set: vi.fn((name: string, value: unknown) => {
      store.set(name, value);
    }),
    remove: vi.fn((name: string) => store.delete(name)),
    saveAsync,
    _store: store,
  };
}

function stubOffice(settings: ReturnType<typeof makeMockSettings> | undefined) {
  vi.stubGlobal('Office', {
    context: settings ? { document: { settings } } : undefined,
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
  });
}

describe('resolveWorkbookId', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('mints and persists a new id when none exists', async () => {
    const settings = makeMockSettings();
    stubOffice(settings);

    const id = await resolveWorkbookId();

    expect(id).toMatch(/^wb_/);
    expect(settings.set).toHaveBeenCalledWith('cellix:workbookId', id);
    expect(settings.saveAsync).toHaveBeenCalled();
  });

  it('returns the existing id without minting a new one', async () => {
    const settings = makeMockSettings({ 'cellix:workbookId': 'wb_existing-123' });
    stubOffice(settings);

    const id = await resolveWorkbookId();

    expect(id).toBe('wb_existing-123');
    expect(settings.set).not.toHaveBeenCalled();
    expect(settings.saveAsync).not.toHaveBeenCalled();
  });

  it('never overwrites an existing value even across repeated calls', async () => {
    const settings = makeMockSettings();
    stubOffice(settings);

    const first = await resolveWorkbookId();
    const second = await resolveWorkbookId();

    expect(second).toBe(first);
    expect(settings.set).toHaveBeenCalledTimes(1);
  });

  it('falls back to a session-local id when document.settings is unavailable', async () => {
    stubOffice(undefined);

    const id = await resolveWorkbookId();

    expect(id).toMatch(/^wb_/);
  });

  it('falls back to a session-local id when Office is not defined at all', async () => {
    vi.unstubAllGlobals();

    const id = await resolveWorkbookId();

    expect(id).toMatch(/^wb_/);
  });

  it('still returns the minted id if saveAsync reports failure (fail-open, per #21)', async () => {
    const settings = makeMockSettings();
    settings.saveAsync.mockImplementation((callback: (result: { status: string; error?: unknown }) => void) => {
      callback({ status: 'failed', error: new Error('mock save failure') });
    });
    stubOffice(settings);

    const id = await resolveWorkbookId();

    expect(id).toMatch(/^wb_/);
    expect(settings.set).toHaveBeenCalledWith('cellix:workbookId', id);
  });
});
