import { ClarificationPayload } from '@/types/cellix.types';

/** Only one clarification question may be pending at a time. */
export function shouldAcceptIncomingClarification(
  pending: ClarificationPayload | null,
): boolean {
  return pending === null;
}

/**
 * Problem solved: The current threshold is too conservative — triggers questions
 * on messages that are 70%+ clear. Users hate being asked questions when the
 * action is obvious.
 *
 * New policy:
 *  - Confidence >= 0.60 → always proceed, no question
 *  - Confidence >= 0.45 + non-destructive → proceed with stated assumption
 *  - Confidence < 0.45 + destructive → one targeted question
 *  - Never ask "which column" — the Planner infers from headers
 *  - Never ask for confirmation — Accept/Reject preview handles that
 */
export interface ClarificationInput {
  /** 0.0–1.0 from LLM Router or Planner */
  confidence: number;
  /** Is the action destructive? (deletes data/sheets) */
  isDestructive: boolean;
  /** Does the message reference a specific column name that exists in headers? */
  columnResolved: boolean;
  /** Does the message reference a specific sheet that exists in the workbook? */
  sheetResolved: boolean;
}

export interface ClarificationDecision {
  shouldAsk: boolean;
  /** If false, state this assumption in the answer */
  assumption?: string;
}

export function shouldAskClarification(input: ClarificationInput): ClarificationDecision {
  const { confidence, isDestructive, columnResolved, sheetResolved } = input;

  if (confidence >= 0.6) {
    return { shouldAsk: false };
  }

  if (confidence >= 0.45 && !isDestructive) {
    return {
      shouldAsk: false,
      assumption: 'I made a best guess based on your sheet — let me know if I got it wrong.',
    };
  }

  if (columnResolved && sheetResolved) {
    return { shouldAsk: false };
  }

  if (confidence < 0.45 && isDestructive) {
    return { shouldAsk: true };
  }

  return { shouldAsk: false };
}

/**
 * Banned clarification questions — these should NEVER be asked.
 * The Planner infers these from sheet context.
 */
export const BANNED_CLARIFICATION_TOPICS = [
  'which column',
  'which sheet',
  'are you sure',
  'do you want to proceed',
  'should i continue',
  'can you confirm',
  'please confirm',
  'do you want me to',
] as const;

export function isBannedClarification(question: string): boolean {
  const lower = question.toLowerCase();
  return BANNED_CLARIFICATION_TOPICS.some((banned) => lower.includes(banned));
}
