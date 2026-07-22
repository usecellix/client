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
  return /\b(as\s+a\s+copy|copy\s+of|duplicate|clone)\b/i.test(message);
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

function tryLocalRenameSheetActions(
  message: string,
  mode: AssistantMode,
): LocalSheetActionPlan | null {
  if (mode !== 'action') return null;
  if (!/rename\s+(?:the\s+)?(?:sheet|tab)/i.test(message)) return null;

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

  return {
    actions: [{ type: 'CLEAR_RANGE', range: 'A1:XFD1048576', mode: 'contents' }],
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
    tryLocalCreateEmptySheetActions(message, workbookContext, mode) ??
    tryLocalRenameSheetActions(message, mode) ??
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

export function extractDeleteSheetNames(
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

  const hasMentions = extractSheetMentions(message).length > 0;
  if (availableSheets.length === 0 && !hasMentions) return null;

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
