// Re-exports the canonical action-type definitions from the repo-root
// `shared/` package. Do not fork a local copy here — see ARCHITECTURE.md AD-7.
// This path has been broken twice: once to `../../Root/shared/action.types`
// (a leftover from when this repo was nested inside a `Root/` directory —
// TASKS.md #175), and reintroduced a second time by a later upstream merge
// with a comment asserting the opposite of the real layout. `Root/` does not
// exist; the package lives at the repo root as `shared/`, i.e. `../../shared`
// from `frontend/src/`. `tsc --noEmit` catches this immediately; vitest does
// not, since it runs without type-checking — that gap is exactly why this
// broke silently both times. TASKS.md #166, #175.
export * from '../../shared/action.types';
