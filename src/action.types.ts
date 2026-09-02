// Re-exports the canonical action-type definitions from the repo-root `shared/`
// package. Do not fork a local copy here — see ARCHITECTURE.md AD-7.
// NOTE: the repo-root package lives at `Root/shared/`. This path was
// `../../shared/action.types` — unresolvable since the rename — so every type
// imported through here silently vanished at compile time. Tests never caught
// it (vitest runs, it does not typecheck), which is exactly the AD-7 blind
// spot this file exists to close. TASKS.md #166.
export * from '../../Root/shared/action.types';
