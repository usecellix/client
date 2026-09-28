// Re-exports the canonical action-type definitions from the local vendored
// `shared/` copy (`frontend/src/shared/action.types.ts`), matching the
// aggregateTable.ts / rangeFilter.ts pattern in the same directory.
//
// This previously re-exported from the repo-root `../../shared/action.types`,
// which only resolves when this package is checked out as a subfolder of the
// monorepo (`usecellix/Root`). This package (`frontend/`) actually ships as
// its own separate repo (`usecellix/client`), which Vercel builds in
// isolation — `../../shared` has never existed in that repo's history, so
// the build failed there even though it always worked locally. Keep this
// vendored copy in sync with `shared/action.types.ts` by hand when either
// changes.
export * from './shared/action.types';
