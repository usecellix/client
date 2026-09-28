import { WorkbookContext } from '@/types/cellix.types';
import { SheetAction } from '@/types/sheet-actions';
import { AssistantMode } from '@/types/mode';
import { extractSheetMentions, stripSheetMentions } from '@/utils/sheetMentions';
import {
  detectCompoundSheetFollowUp,
  extractSheetNameFromPrompt,
  sanitizeExcelSheetName,
  SHEET_NAME_BOUNDARY,
} from '@/utils/sheetName.util';

export interface LocalSheetActionPlan {
  actions: SheetAction[];
  explanation: string;
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function detectCreateSheetIntent(message: string): boolean {
  return /\b(create|add)\s+(?:an?\s+)?(?:(?:new|empty|blank)\s+)*(?:sheet|tab)\b/i.test(message);
}

export function detectCopySheetIntent(message: string): boolean {
  // Mirrors the server's compound-action.util.ts — "Copy the Purchase Register
  // sheet and name it March Copy" must never look like a blank-sheet create.
  // TASKS.md #213.
  return (
    /\b(as\s+a\s+copy|copy\s+of|duplicate|clone|replicate)\b/i.test(message) ||
    /\bcopy\s+(?:[A-Za-z0-9_'-]+\s+){0,4}?(?:sheet|tab)\b/i.test(message) ||
    /\bcopy\s+(?:sheet|tab)\b/i.test(message)
  );
}

export function detectSortIntent(message: string): boolean {
  return (
    /\bsort(?:\s+the\s+values?\s+of|\s+(?:the\s+)?(?:sheet\s+)?(?:based\s+on|by|on)|\s+based\s+on|\s+by|\s+on|\s+column\b)/i.test(
      message,
    ) || /\bin\s+(?:ascending|descending)\s+order\b/i.test(message)
  );
}

/** Prompts that need LLM planning (data, copy, sort, etc.) — not empty-sheet-only. */
export function detectSheetDataGenerationIntent(message: string): boolean {
  const lower = message.toLowerCase();
  if (detectCompoundSheetFollowUp(message)) return true;
  if (detectCopySheetIntent(message) || detectSortIntent(message)) return true;

  if (/\badd\s+a\s+total\b/i.test(message) || /\btotal\s+row\b/i.test(message)) return true;
  if (/\bsample\s+rows?\b/i.test(message)) return true;
  if (/\bfill\s+with\b/i.test(message)) return true;
  if (/\bwith\s+headers?\b/i.test(message) && /\b(create|add|make|build)\b/i.test(message)) {
    return true;
  }

  const hasDataKeyword =
    /\b(dummy|sample|data|values?|rows?|headers?|columns?|populate|generate|fill|table|content|gst)\b/i.test(
      lower,
    );
  const hasCreateKeyword = /\b(create|add|generate|populate|fill|make|build|give|insert)\b/i.test(
    lower,
  );

  if (/\bchart\b/i.test(message) && /\banaly(?:sis|ze|se)\b/i.test(message)) return true;

  return hasDataKeyword && hasCreateKeyword;
}

export function detectEmptySheetCreateIntent(message: string): boolean {
  if (!detectCreateSheetIntent(message)) return false;
  if (detectSheetDataGenerationIntent(message)) return false;
  return true;
}

export { extractSheetNameFromPrompt } from '@/utils/sheetName.util';

function nextUniqueSheetName(base: string, availableSheets: string[]): string {
  const existing = new Set(availableSheets.map((sheet) => sheet.toLowerCase()));
  if (!existing.has(base.toLowerCase())) return base;
  for (let i = 2; i <= 99; i += 1) {
    const candidate = `${base} ${i}`;
    if (!existing.has(candidate.toLowerCase())) return candidate;
  }
  return `${base} ${Date.now()}`;
}

export function suggestNewSheetName(message: string, availableSheets: string[]): string {
  const fromPrompt = extractSheetNameFromPrompt(message);
  if (fromPrompt) return nextUniqueSheetName(fromPrompt, availableSheets);

  if (/\bsorted\b/i.test(message)) return nextUniqueSheetName('Sorted', availableSheets);
  return nextUniqueSheetName('Sheet2', availableSheets);
}

export function extractCreateSheetName(
  message: string,
  availableSheets: string[],
): string | null {
  const fromPrompt = extractSheetNameFromPrompt(message);
  if (fromPrompt) return nextUniqueSheetName(fromPrompt, availableSheets);

  const quoted = extractQuotedNames(message);
  if (quoted.length > 0) {
    const resolved = resolveSheetNames(quoted, availableSheets);
    if (resolved[0]) return resolved[0];
  }

  const afterSheet = new RegExp(
    `\\b(?:create|add)\\s+(?:an?\\s+)?(?:(?:new|empty|blank)\\s+)*sheets?\\s+(?:named\\s+|called\\s+)?([A-Za-z][A-Za-z0-9 _-]+?)(?=${SHEET_NAME_BOUNDARY})`,
    'i',
  ).exec(message);
  const candidate = afterSheet?.[1]?.trim();
  if (candidate && !/^(with|and|named|called)$/i.test(candidate)) {
    return nextUniqueSheetName(candidate, availableSheets);
  }

  const sortedSheets = [...availableSheets].sort((a, b) => b.length - a.length);
  for (const sheet of sortedSheets) {
    if (new RegExp(`\\b${escapeRegex(sheet)}\\b`, 'i').test(message)) {
      return sheet;
    }
  }

  return suggestNewSheetName(message, availableSheets);
}

export function buildCreateEmptySheetActions(sheetName: string): SheetAction[] {
  return [{ type: 'ADD_SHEET', name: sanitizeExcelSheetName(sheetName) }];
}

export function buildCreateEmptySheetExplanation(sheetName: string): string {
  return `Create empty sheet "${sheetName}"`;
}

export function tryLocalCreateEmptySheetActions(
  message: string,
  workbookContext: WorkbookContext | undefined,
  mode: AssistantMode,
): LocalSheetActionPlan | null {
  if (mode !== 'action') return null;
  if (!detectEmptySheetCreateIntent(message)) return null;

  const availableSheets = (workbookContext?.sheets ?? [])
    .map((sheet) => sheet.sheetName)
    .filter(Boolean);

  const sheetName = extractCreateSheetName(message, availableSheets);
  if (!sheetName) return null;

  return {
    actions: buildCreateEmptySheetActions(sheetName),
    explanation: buildCreateEmptySheetExplanation(sheetName),
  };
}

/**
 * "Copy the Purchase Register sheet and name it March Copy" — the guide's own
 * T1.1 phrasing — used to always go through the full Tier 3 planner/executor/
 * verifier LLM pipeline (5-40s observed live) for what is a fully
 * deterministic operation: extract the source and destination sheet names,
 * confirm the source exists, emit ONE ADD_SHEET{copyFrom} action. Conservative
 * on purpose — only the clear "copy X sheet and name/call it Y" shape matches;
 * anything else (a filtered/partial copy, an unresolvable source name, no
 * match at all) falls through to the backend exactly as before.
 */
function tryLocalCopySheetActions(
  message: string,
  workbookContext: WorkbookContext | undefined,
  mode: AssistantMode,
): LocalSheetActionPlan | null {
  if (mode !== 'action') return null;
  if (!detectCopySheetIntent(message)) return null;

  const availableSheets = (workbookContext?.sheets ?? [])
    .map((sheet) => sheet.sheetName)
    .filter(Boolean);
  if (availableSheets.length === 0) return null;

  const match = message.match(
    /\bcopy\s+(?:the\s+)?["']?([^"'\n]+?)["']?\s+(?:sheet|tab)\b[\s\S]{0,20}?\b(?:and\s+)?(?:name|call)\s+it\s+["']?([^"'\n]+?)["']?\s*[.!]?\s*$/i,
  );
  if (!match) return null;

  const sourceCandidate = match[1]?.trim();
  const destCandidate = match[2]?.trim();
  if (!sourceCandidate || !destCandidate) return null;

  const [sourceName] = resolveSheetNames([sourceCandidate], availableSheets);
  // Source doesn't resolve to a real sheet — let the backend interpret/clarify
  // rather than guess and risk copying the wrong sheet.
  if (!sourceName) return null;

  const newSheetName = sanitizeExcelSheetName(nextUniqueSheetName(destCandidate, availableSheets));

  return {
    actions: [{ type: 'ADD_SHEET', name: newSheetName, copyFrom: sourceName }],
    explanation: `Copy sheet "${sourceName}" to "${newSheetName}"`,
  };
}

function tryLocalRenameSheetActions(
  message: string,
  workbookContext: WorkbookContext | undefined,
  mode: AssistantMode,
): LocalSheetActionPlan | null {
  if (mode !== 'action') return null;
  if (!/rename\s+(?:the\s+)?(?:sheet|tab|this|current|active)/i.test(message)) return null;

  // "Rename this sheet to X" / "rename the current tab to X" / "rename the
  // active sheet to X" — the guide's own canonical T1.1 phrasing, referring
  // to whatever sheet is currently active rather than naming one. Missing
  // this sent every such request through the full LLM pipeline (5-6s
  // observed live) instead of the same instant local lane a named rename
  // already gets. TASKS.md #254.
  const activeSheetMatch = message.match(
    /rename\s+(?:the\s+)?(?:this|current|active)\s+(?:sheet|tab)\s+to\s+["']?([^"']+?)["']?\s*[.!]?\s*$/i,
  );
  if (activeSheetMatch) {
    const oldName = workbookContext?.activeSheet;
    const newName = activeSheetMatch[1]?.trim();
    if (!oldName || !newName) return null;
    return {
      actions: [{ type: 'RENAME_SHEET', oldName, newName }],
      explanation: `Rename sheet "${oldName}" to "${newName}"`,
    };
  }

  const match = message.match(
    /rename\s+(?:the\s+)?(?:sheet|tab)\s+["']?([^"']+?)["']?\s+to\s+["']?([^"']+?)["']?\s*$/i,
  );
  if (!match) return null;

  const oldName = match[1].trim();
  const newName = match[2].trim();
  return {
    actions: [{ type: 'RENAME_SHEET', oldName, newName }],
    explanation: `Rename sheet "${oldName}" to "${newName}"`,
  };
}

/**
 * Words that survive stripping the clear phrase itself and still mean "the
 * whole active sheet" — anything else left over is a narrower target.
 */
const WHOLE_SHEET_CLEAR_FILLER =
  /^(?:\s|[.,!?]|\b(?:clear|out|this|that|the|entire|whole|complete|completely|all|of|please|now|just|kindly|sheet|tab|worksheet|data|content|contents|cell|cells|value|values|everything|active|current|open)\b)*$/i;

/**
 * The clear trigger matches "clear all (the) data/content/cells" anywhere in the
 * message, so scoped clears — "clear all data in column C", "clear all the
 * content in the Narration column", "clear all cells with errors", "clear the
 * Summary sheet" — were answered with CLEAR_RANGE A1:XFD1048576 on the ACTIVE
 * sheet: the entire sheet wiped (plus its charts, #181) for a column-sized
 * request, or the wrong sheet entirely. A whole-sheet clear is only a
 * whole-sheet clear when nothing in the message narrows it; everything else
 * goes to the backend, which can resolve columns, conditions and sheet names.
 * TASKS.md #209.
 */
export function isWholeSheetClear(message: string): boolean {
  const withoutMentions = stripSheetMentions(message);
  return WHOLE_SHEET_CLEAR_FILLER.test(withoutMentions.trim());
}

function tryLocalClearSheetActions(
  message: string,
  mode: AssistantMode,
): LocalSheetActionPlan | null {
  if (mode !== 'action') return null;
  if (
    !/clear\s+(?:this\s+)?(?:entire\s+)?(?:sheet|all\s+(?:the\s+)?(?:data|content|cells))/i.test(
      message,
    )
  ) {
    return null;
  }
  if (/[A-Z]+\d+:[A-Z]+\d+/i.test(message)) return null;
  if (!isWholeSheetClear(message)) return null;

  return {
    // Whole-sheet clear means "make it a plain workbook" — cell contents AND
    // any charts left floating over the (now empty) grid. TASKS.md #181.
    actions: [
      { type: 'CLEAR_RANGE', range: 'A1:XFD1048576', mode: 'contents', clearCharts: true },
    ],
    explanation: 'Clear all data on the active sheet',
  };
}

export function tryLocalSheetActions(
  message: string,
  workbookContext: WorkbookContext | undefined,
  mode: AssistantMode,
): LocalSheetActionPlan | null {
  return (
    tryLocalDeleteSheetActions(message, workbookContext, mode) ??
    tryLocalCopySheetActions(message, workbookContext, mode) ??
    tryLocalCreateEmptySheetActions(message, workbookContext, mode) ??
    tryLocalRenameSheetActions(message, workbookContext, mode) ??
    tryLocalClearSheetActions(message, mode) ??
    null
  );
}

export function detectDeleteSheetIntent(message: string): boolean {
  const lower = message.toLowerCase();
  return (
    /\b(delete|remove|drop)\b/.test(lower) &&
    (/\bsheets?\b/.test(lower) || /\btab(s)?\b/.test(lower))
  );
}

/**
 * Words that, as the object of the delete verb, mean the user is deleting
 * something *inside* a sheet rather than the sheet itself.
 */
const NON_SHEET_DELETE_OBJECT =
  /\b(rows?|columns?|cols?|cells?|duplicates?|dupes|blanks?|values?|entries|entry|records?|data|contents?|formatting|formats?|formulas?|comments?|notes?|charts?|graphs?|tables?|filters?|borders?|colou?rs?|text|spaces?|errors?|headers?|totals?|lines?|items?|everything|anything)\b/i;

/** "in/from/on the X sheet" places the sheet as a location, not the thing being deleted. */
const SHEET_AS_LOCATION = /\b(in|from|on|within|inside|across)\b/i;

/**
 * detectDeleteSheetIntent only checks that a delete verb and the word
 * "sheet"/"tab" both appear, so "Delete blank rows in the Summary sheet" and
 * "Remove duplicates from this sheet" looked like sheet deletes and this lane
 * proposed DELETE_SHEET before the backend was ever asked. This checks that the
 * sheet is the verb's object: the words between the verb and "sheet"/"tab"
 * (with real sheet names and @[mentions] removed, so a sheet called "Rows Data"
 * still works) must not name something else or use the sheet as a location.
 * Mirrors the server's local-sheet-actions.util.ts. TASKS.md #208.
 */
export function isSheetTheDeleteObject(message: string, availableSheets: string[]): boolean {
  const match = /\b(?:delete|remove|drop)\b([\s\S]*?)\b(?:sheets?|tabs?)\b/i.exec(message);
  if (!match) return false;

  let objectPhrase = stripSheetMentions(match[1]);
  const sortedSheets = [...availableSheets].sort((a, b) => b.length - a.length);
  for (const sheet of sortedSheets) {
    objectPhrase = objectPhrase.replace(new RegExp(`\\b${escapeRegex(sheet)}\\b`, 'gi'), ' ');
  }

  return !NON_SHEET_DELETE_OBJECT.test(objectPhrase) && !SHEET_AS_LOCATION.test(objectPhrase);
}

/**
 * "Delete all sheets except X" / "keep X remove the rest" — named sheets are preserves, not targets.
 */
export function isPreserveOtherSheetsDelete(message: string): boolean {
  const lower = message.toLowerCase();
  if (/\b(except|except for|apart from|other than)\b/.test(lower)) return true;
  if (
    /\b(keep|keeping|leave|leaving)\b/.test(lower) &&
    /\b(all|other|every|rest|remaining)\b/.test(lower)
  ) {
    return true;
  }
  if (
    /\b(delete|remove|drop)\s+(all|every)\b/.test(lower) &&
    /\b(keep|except|but)\b/.test(lower)
  ) {
    return true;
  }
  return false;
}

function extractQuotedNames(message: string): string[] {
  const names: string[] = [];
  const pattern = /["']([^"']+)["']/g;
  let match = pattern.exec(message);
  while (match) {
    const name = match[1]?.trim();
    if (name) names.push(name);
    match = pattern.exec(message);
  }
  return names;
}

function resolveSheetNames(candidates: string[], availableSheets: string[]): string[] {
  const byLower = new Map(availableSheets.map((sheet) => [sheet.toLowerCase(), sheet]));
  const resolved: string[] = [];

  for (const candidate of candidates) {
    const trimmed = candidate.trim();
    if (!trimmed) continue;
    const exact = byLower.get(trimmed.toLowerCase());
    if (exact && !resolved.includes(exact)) {
      resolved.push(exact);
    }
  }

  return resolved;
}

/** Sheets named in the message (mentions, quotes, active-sheet phrase, or workbook name scan). */
export function extractReferencedSheetNames(
  message: string,
  availableSheets: string[],
  activeSheet?: string,
): string[] {
  const mentions = extractSheetMentions(message);
  if (mentions.length > 0) {
    const resolved = resolveSheetNames(mentions, availableSheets);
    if (resolved.length > 0) return resolved;
    return mentions;
  }

  if (/\b(this|current|active)\s+sheet\b/i.test(message) && activeSheet) {
    return [activeSheet];
  }

  const quoted = extractQuotedNames(message);
  if (quoted.length > 0) {
    const resolved = resolveSheetNames(quoted, availableSheets);
    if (resolved.length > 0) return resolved;
    return quoted;
  }

  const cleaned = stripSheetMentions(message);
  const sortedSheets = [...availableSheets].sort((a, b) => b.length - a.length);
  const mentioned: string[] = [];
  for (const sheet of sortedSheets) {
    const pattern = new RegExp(`\\b${escapeRegex(sheet)}\\b`, 'i');
    if (pattern.test(cleaned) && !mentioned.includes(sheet)) {
      mentioned.push(sheet);
    }
  }
  if (mentioned.length > 0) {
    mentioned.sort(
      (a, b) =>
        cleaned.toLowerCase().indexOf(a.toLowerCase()) -
        cleaned.toLowerCase().indexOf(b.toLowerCase()),
    );
    return mentioned;
  }

  const exceptClause =
    /\b(?:except(?:\s+for)?|apart\s+from|other\s+than|keep(?:ing)?|leave(?:ing)?)\s+(.+?)(?:[.!?]|$)/i.exec(
      cleaned,
    );
  if (exceptClause?.[1]) {
    const part = exceptClause[1]
      .replace(/\b(?:sheet|tab)s?\b/gi, ' ')
      .replace(/\b(?:named|called)\b/gi, ' ')
      .trim();
    const parts = part
      .split(/\s*,\s*|\s+and\s+/i)
      .map((p) => p.replace(/^["']|["']$/g, '').trim())
      .filter(Boolean);
    const resolved = resolveSheetNames(parts, availableSheets);
    if (resolved.length > 0) return resolved;
  }

  return [];
}

export function extractDeleteSheetNames(
  message: string,
  availableSheets: string[],
  activeSheet?: string,
): string[] {
  if (isPreserveOtherSheetsDelete(message)) {
    if (availableSheets.length === 0) return [];
    const preserve = extractReferencedSheetNames(message, availableSheets, activeSheet);
    if (preserve.length === 0) {
      return [];
    }
    const preserveSet = new Set(preserve.map((name) => name.toLowerCase()));
    return availableSheets.filter((name) => !preserveSet.has(name.toLowerCase()));
  }

  const referenced = extractReferencedSheetNames(message, availableSheets, activeSheet);
  if (referenced.length > 0) return referenced;

  const cleaned = stripSheetMentions(message);
  const listMatch =
    /\b(?:delete|remove|drop)\s+(?:the\s+)?sheets?\s+(?:named\s+)?(.+?)(?:[.!?]|$)/i.exec(
      cleaned,
    );
  if (listMatch?.[1]) {
    const parts = listMatch[1]
      .split(/\s*,\s*|\s+and\s+/i)
      .map((part) => part.replace(/^["']|["']$/g, '').trim())
      .filter(Boolean);
    const resolved = resolveSheetNames(parts, availableSheets);
    if (resolved.length > 0) return resolved;
    return parts;
  }

  return [];
}

export function buildDeleteSheetActions(sheetNames: string[]): SheetAction[] {
  return sheetNames.map((sheetName) => ({
    type: 'DELETE_SHEET' as const,
    sheetName,
  }));
}

export function buildDeleteSheetExplanation(sheetNames: string[]): string {
  if (sheetNames.length === 1) {
    return `Delete sheet "${sheetNames[0]}"`;
  }
  return `Delete sheets: ${sheetNames.map((name) => `"${name}"`).join(', ')}`;
}

export function tryLocalDeleteSheetActions(
  message: string,
  workbookContext: WorkbookContext | undefined,
  mode: AssistantMode,
): LocalSheetActionPlan | null {
  if (mode !== 'action') return null;
  if (!detectDeleteSheetIntent(message)) return null;

  const availableSheets = (workbookContext?.sheets ?? [])
    .map((sheet) => sheet.sheetName)
    .filter(Boolean);
  if (!isSheetTheDeleteObject(message, availableSheets)) return null;

  const hasMentions = extractSheetMentions(message).length > 0;
  if (availableSheets.length === 0 && !hasMentions) return null;

  if (isPreserveOtherSheetsDelete(message) && availableSheets.length === 0) {
    return null;
  }

  const sheetNames = extractDeleteSheetNames(
    message,
    availableSheets,
    workbookContext?.activeSheet,
  );

  if (sheetNames.length === 0) return null;

  const actions = buildDeleteSheetActions(sheetNames);
  return {
    actions,
    explanation: buildDeleteSheetExplanation(sheetNames),
  };
}
