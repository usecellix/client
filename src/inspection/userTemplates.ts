/**
 * P3.24 — firm template library, user part. A saved template is a NAME + HEADER ROW only
 * (never cell values), kept in this browser's storage as a per-viewer convenience.
 * Storage can be blocked (private window, cleared data) — every access is guarded.
 */
export interface UserTemplate {
  name: string;
  headers: string[];
}

const KEY = 'cellix.userTemplates.v1';

export function listUserTemplates(): UserTemplate[] {
  try {
    const raw = globalThis.localStorage?.getItem(KEY);
    const parsed = raw ? (JSON.parse(raw) as unknown) : [];
    return Array.isArray(parsed)
      ? parsed.filter((t): t is UserTemplate => !!t && typeof t.name === 'string' && Array.isArray(t.headers))
      : [];
  } catch {
    return [];
  }
}

export function saveUserTemplate(t: UserTemplate): boolean {
  const clean: UserTemplate = {
    name: t.name.trim().slice(0, 31),
    // Headers only — strings, trimmed, no values ever.
    headers: t.headers.map((h) => String(h).trim()).filter(Boolean).slice(0, 100),
  };
  if (!clean.name || !clean.headers.length) return false;
  try {
    const rest = listUserTemplates().filter((x) => x.name.toLowerCase() !== clean.name.toLowerCase());
    globalThis.localStorage?.setItem(KEY, JSON.stringify([...rest, clean]));
    return true;
  } catch {
    return false;
  }
}
