// Re-exports the canonical range-filter logic from the repo-root `shared/`
// package. Do not fork a local copy here — see TASKS.md #61.
export {
  applyFilterOperator,
  buildOutputRows,
  filterDataRows,
  findMatchingRowOffsets,
  resolveFilterColumnIndex,
  type RangeFilterOperator,
  type RangeFilterSpec,
} from '../../../shared/rangeFilter';
