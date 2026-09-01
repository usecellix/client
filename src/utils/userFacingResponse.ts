import { CellChange } from '@/types/changeSet';
import { SheetAction } from '@/types/sheet-actions';

export interface UserFacingSummary {
  contextLine?: string;
  headline: string;
  bullets?: string[];
  supportingDetail?: string;
}

export interface ResponseInternalDetails {
  tier?: number;
  model?: string;
  processingLabel?: string;
  reasoning?: string;
  assumption?: string;
  rawActionSummary?: string;
  legacyExplanation?: string;
}

export const INTERNAL_COPY_MARKERS =
  /\b(Tier\s*[0-3]|single-action|no verification|Direct Change|Planner|Executor|Verifier|CONDITIONAL_FORMAT|FORMAT_MATCHING_ROWS|findMatchingRowOffsets|hasHeaders\s*:|SET_FORMULA|WRITE_TABLE|openai\/)/i;

function looksInternal(text: string): boolean {
  return INTERNAL_COPY_MARKERS.test(text);
}

function stripMarkdown(text: string): string {
  return text
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\*([^*]+)\*/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .trim();
}

/** Resolve display copy for an action card (prefers backend userFacingSummary). */
export function resolveActionBlockCopy(input: {
  userFacingSummary?: UserFacingSummary;
  explanation?: string;
  actions: SheetAction[];
  changes?: CellChange[];
}): UserFacingSummary {
  if (input.userFacingSummary?.headline) {
    return input.userFacingSummary;
  }

  const changes = input.changes ?? [];
  let headline = stripMarkdown(input.explanation ?? '');
  if (!headline || looksInternal(headline)) {
    const n = changes.length || input.actions.length;
    headline =
      n <= 0
        ? 'Ready to apply changes.'
        : n === 1
          ? "I'll apply 1 change to your sheet."
          : `I'll apply ${n} changes to your sheet.`;
  } else if (!/[.!?]$/.test(headline)) {
    headline = `${headline}.`;
  }

  let supportingDetail: string | undefined;
  if (changes.length > 0) {
    supportingDetail = `${changes.length} cell${changes.length === 1 ? '' : 's'}`;
  }

  return { headline, supportingDetail };
}

export function hasInternalDetails(details?: ResponseInternalDetails): boolean {
  if (!details) return false;
  return Boolean(
    details.model ||
      details.processingLabel ||
      details.reasoning ||
      details.assumption ||
      details.rawActionSummary ||
      details.legacyExplanation ||
      typeof details.tier === 'number',
  );
}

export function formatInternalDetailsLines(details: ResponseInternalDetails): string[] {
  const lines: string[] = [];
  if (details.model) lines.push(`Model: ${details.model}`);
  if (typeof details.tier === 'number' || details.processingLabel) {
    lines.push(
      `Processing: ${details.processingLabel ?? `Tier ${details.tier}`}`,
    );
  }
  if (details.assumption) lines.push(`Assumption: ${details.assumption}`);
  if (details.reasoning) lines.push(`Reasoning: ${details.reasoning}`);
  if (details.rawActionSummary) lines.push(`Raw action: ${details.rawActionSummary}`);
  if (details.legacyExplanation && details.legacyExplanation !== details.processingLabel) {
    lines.push(details.legacyExplanation);
  }
  return lines;
}
