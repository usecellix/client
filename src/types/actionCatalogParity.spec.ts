import { describe, expect, it } from 'vitest';
import { ALL_FRONTEND_SHEET_ACTION_TYPES } from './sheetActionCatalog';
// Cross-package import into the backend's own exhaustiveness catalog — see
// TASKS.md #5. This is deliberately the backend's *runtime* export
// (Object.keys of its Record<SheetActionType, CatalogEntry>), not a
// hand-copied list, so it can never itself drift from what the backend
// actually declares.
import { ALL_SHEET_ACTION_TYPES } from '../../../Server/src/excel-ai/types/action-catalog';

// Compared as plain strings, not the two packages' distinct SheetActionType
// literal unions — those two types are legitimately different (frontend's
// includes two rich-only extras, see below), so forcing one array's element
// type onto the other doesn't type-check. The exhaustiveness guarantee lives
// in each catalog's own Record<X, true>, not in this comparison.
const backendTypes = new Set<string>(ALL_SHEET_ACTION_TYPES);
const frontendTypes = new Set<string>(ALL_FRONTEND_SHEET_ACTION_TYPES);

describe('action catalog parity — frontend vs backend (TASKS.md #5)', () => {
  it('recognizes every wire-level action type the backend catalog knows about', () => {
    // Not full set equality: the backend catalog is allowed extra
    // internal-only concepts (it isn't — today they're 1:1 — but nothing
    // requires it). What must never happen is the backend advertising or
    // emitting a type the frontend has no declared knowledge of at all,
    // which is exactly the "type exists in the schema, unhandled on the
    // other side" bug class this whole check exists to catch (see spec 14 /
    // the original FREEZE_PANES incident that motivated the backend's own
    // action-catalog.ts).
    const missing = [...backendTypes].filter((type) => !frontendTypes.has(type));
    expect(missing).toEqual([]);
  });

  it('does not carry frontend-only types the backend has never heard of, without a documented reason', () => {
    // The two known, intentional exceptions: CLEAR_RANGE and APPEND_ROW are
    // rich-only forms produced by legacyConverter.ts and the local shortcut
    // lane — they never arrive as a raw wire type from the backend. Anything
    // else here that the backend doesn't know about is worth a second look
    // rather than silent acceptance.
    const KNOWN_FRONTEND_ONLY = new Set(['CLEAR_RANGE', 'APPEND_ROW']);
    const frontendOnly = [...frontendTypes].filter(
      (type) => !backendTypes.has(type) && !KNOWN_FRONTEND_ONLY.has(type),
    );
    expect(frontendOnly).toEqual([]);
  });
});
