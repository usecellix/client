/**
 * Compact copy for action previews — models sometimes dump cell-by-cell inventories.
 * Keep the UI to a short summary; Accept/Reject already shows the change count.
 */
const DETAIL_MARKERS =
  /\b(Exactly what will change|Cells affected|Values to be written|Expected result|No formulas will be used)\b/i;

const ROW_DUMP = /\bRow\s+\d+\s*->/i;

export function shortenActionPreviewCopy(text: string, maxLen = 160): string {
  const raw = text.trim();
  if (!raw) return '';

  const cutAtMarker = raw.search(DETAIL_MARKERS);
  let compact = cutAtMarker > 0 ? raw.slice(0, cutAtMarker).trim() : raw;

  const previewMatch = compact.match(/^Preview\s*\(([^)]+)\)\s*:?\s*/i);
  if (previewMatch) {
    // Prefer the short intent inside "Preview (...)" and drop row dumps after it.
    compact = previewMatch[1].trim();
  } else {
    const dumpAt = compact.search(ROW_DUMP);
    if (dumpAt > 0) {
      compact = compact.slice(0, dumpAt).trim().replace(/[:\-–—]\s*$/, '');
    }
    compact = (compact.split(/\r?\n/)[0] ?? compact).trim();
  }

  // Drop long header lists: "under headers A | B | C | ..."
  compact = compact.replace(/\bunder headers?\b[\s\S]*/i, 'under the existing headers').trim();

  if (compact.length <= maxLen) {
    return compact.endsWith('.') ? compact : `${compact}.`;
  }

  const truncated = compact.slice(0, maxLen - 1);
  const lastSpace = truncated.lastIndexOf(' ');
  return `${(lastSpace > 40 ? truncated.slice(0, lastSpace) : truncated).trimEnd()}…`;
}

export function summarizeActionsFallback(actionCount: number): string {
  if (actionCount <= 0) return 'Ready to apply changes.';
  if (actionCount === 1) return '1 change ready to apply.';
  return `${actionCount} changes ready to apply.`;
}
