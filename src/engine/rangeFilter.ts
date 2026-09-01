// Re-exports the canonical range-filter logic from the local `shared/`
// package (client/src/shared/), matching the aggregateTable.ts pattern —
// the repo-root shared/ package used elsewhere in this file's history is
// not part of this repo checkout.
export {
  applyFilterOperator,
  buildOutputRows,
  filterDataRows,
  findMatchingRowOffsets,
  resolveFilterColumnIndex,
  type RangeFilterOperator,
  type RangeFilterSpec,
} from '../shared/rangeFilter';
