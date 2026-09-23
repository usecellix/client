/**
 * Spec 24 — never put raw engine/validation errors into chat UI.
 */

const APPLY_ERROR_FALLBACK =
  "I couldn't apply that formatting. Please try again or describe the range differently.";

const INTERNAL_APPLY_ERROR_RE =
  /findMatchingRowOffsets|hasHeaders\s*:|FORMAT_MATCHING_ROWS|SET_MATCHING_ROWS|WRITE_TABLE|Spreadsheet update failed|Unsupported action|RichActionEngine|virtualApply|columnIndex|RangeCopyType|Office\.js|TypeError|ReferenceError|is not a function|at\s+\S+\s+\(/i;

/** ActionType: message  or  ActionType: message; ActionType2: ... */
const ACTION_TYPE_PREFIX_RE = /\b[A-Z][A-Z0-9_]+\s*:\s*/;

/**
 * Map apply/engine errors to a clean user-facing string.
 * Full technical detail should stay in console logs only.
 */
export function toUserFacingApplyError(message: string): string {
  const raw = (message ?? '').trim();
  if (!raw) {
    return "I couldn't apply those changes. Please try again.";
  }

  // A missing sheet means two different things depending on when it is hit,
  // and the action-type prefix is what tells them apart — TASKS.md #263.
  //
  // The engine throws `${action.type}: ${message}` at APPLY time, so an
  // ItemNotFound carrying a prefix means a sheet this change set was supposed
  // to create does not exist — telling the user to click Accept is useless,
  // they just did. Without a prefix it is the PREVIEW path, where the sheet
  // legitimately does not exist yet and Accept is exactly the right advice
  // (handled further down, unchanged).
  //
  // Before this, the prefixed case fell into ACTION_TYPE_PREFIX_RE below and
  // reported a formatting problem — live, a BATCH_SET onto a Main sheet no
  // step had created told the user to re-describe a range that was never the
  // issue.
  const isItemNotFound = /requested resource doesn'?t exist|itemnotfound/i.test(raw);
  if (isItemNotFound && ACTION_TYPE_PREFIX_RE.test(raw)) {
    return (
      "Excel couldn't find a sheet one of those steps needed — it was never created. " +
      'Re-run the request so the missing sheet gets created first.'
    );
  }

  if (INTERNAL_APPLY_ERROR_RE.test(raw) || ACTION_TYPE_PREFIX_RE.test(raw)) {
    return APPLY_ERROR_FALLBACK;
  }

  // Overwrite guard messages are intentionally user-facing.
  if (/write blocked|overwrite|occupied/i.test(raw)) {
    return raw.length > 280 ? `${raw.slice(0, 277)}…` : raw;
  }

  // Generic long/stacky payloads
  if (raw.includes('\n    at ') || raw.length > 200) {
    return "I couldn't apply those changes. Please try again.";
  }

  // Excel host: sheet/range not found (often new sheets not soft-previewed yet).
  if (/requested resource doesn'?t exist|itemnotfound|workbook\.worksheets/i.test(raw)) {
    return (
      "Excel couldn't open a sheet or range for preview (it may only exist after Accept). " +
      "Your change package is still listed — click Accept to apply the create/write steps."
    );
  }

  return raw.length > 280 ? `${raw.slice(0, 277)}…` : raw;
}
