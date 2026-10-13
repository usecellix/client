import type { RepairRequest } from '@/services/repairRequest';
import { AssistantMode } from '@/types/mode';
import type { SpillBlockage } from '@/services/spillBlockers';
import { SheetAction } from '@/hooks/useSseStream';
import { CellChange } from '@/types/changeSet';
import type {
  ResponseInternalDetails,
  UserFacingSummary,
} from '@/utils/userFacingResponse';

export type StepPhase = 'hidden' | 'revealed' | 'running' | 'done';

export interface StepBlock {
  id: string;
  type: 'step';
  label: string;
  phase: StepPhase;
}

/** @deprecated use phase */
export type BlockStatus = 'active' | 'done';

export interface ThinkingBlock {
  id: string;
  type: 'thinking';
  content: string;
  expanded: boolean;
  loading?: boolean;
  visible?: boolean;
}

export interface AnswerBlock {
  id: string;
  type: 'answer';
  content: string;
  revealState: 'hidden' | 'typing' | 'complete';
  matches?: MatchResult[];
}

export interface QuestionBlock {
  id: string;
  type: 'question';
  question: string;
  options?: string[];
  revealState?: 'hidden' | 'visible';
  /**
   * What the user replied. Set when the question is answered, which resolves
   * the block in place inside the turn that asked it rather than the answer
   * appearing as a separate user turn further down the thread. TASKS.md #194.
   */
  answeredWith?: string;
}

export interface ActionBlock {
  id: string;
  type: 'actions';
  actions: SheetAction[];
  explanation: string;
  proposalStatus: 'pending' | 'accepted' | 'rejected';
  changeSetId?: string;
  changes?: CellChange[];
  userFacingSummary?: UserFacingSummary;
  internalDetails?: ResponseInternalDetails;
  /**
   * Staged accept waves (large multi-sheet builds split "create the sheets"
   * from "fill them in"): when set, this block must not be accepted until the
   * sibling block whose changeSetId matches this value is 'accepted' — its
   * actions (e.g. sheet creates) must actually exist first.
   */
  dependsOnChangeSetId?: string;
  /**
   * Position within a staged build (TASKS.md #160). Present only when the
   * server split the work into steps; absent for a single-card change.
   */
  stepIndex?: number;
  stepTotal?: number;
  stepLabel?: string;
  /**
   * Step-wise run this card belongs to (TASKS.md #153). Present only when the
   * run is PAUSED on this card: the backend has generated nothing beyond it,
   * and deciding this card is what triggers the next wave's generation.
   */
  runId?: string;
  stepwise?: boolean;
  /**
   * Action types in this batch with no defined inverse (per the backend's
   * reversibility-catalog.ts). Surfaced here so the user is warned before
   * Accept, not only discovered later when a revert fails.
   */
  irreversibleActionTypes?: string[];
  /**
   * Set once this accepted change has been reverted. Kept on the block, not in
   * the revert control, so the card and the menu still say so after the menu
   * is closed or the pane is reloaded. TASKS.md #404.
   */
  reverted?: boolean;
  /**
   * Set only for a GST-recon "missed rows" batch — the sheet name its
   * CREATE_SHEET/WRITE_TABLE actions target. Lets Accept do one last, fresh
   * existence check for that specific name immediately before applying (the
   * name was free when this card was built, but the workbook could have
   * changed since) and, if it now exists, redirect to the Overwrite/Create-new
   * collision card instead of writing over it silently.
   */
  gstReconOutputSheetName?: string;
  /**
   * Set only for an attachment import (bank statement): the new sheet its
   * actions create. Accept re-checks that the name is still free and, if a
   * sheet has taken it since the card was built, moves the import to the next
   * free name. An import never writes into an existing sheet.
   *
   * Also marks the block for `chatSessionStorage`, which keeps the imported
   * rows out of the copy it saves.
   */
  importSheetName?: string;
  /**
   * The question the user typed alongside an attached statement. It is asked
   * as an ordinary chat turn once this import has been accepted, because only
   * then is the statement in the workbook for the assistant to read.
   */
  followUp?: { prompt: string; mode: AssistantMode };
}

export interface StatusBlock {
  id: string;
  type: 'status';
  label: string;
  pulsing: boolean;
  visible?: boolean;
}

export interface PlanStep {
  title: string;
  detail?: string;
}

export interface PlanBlock {
  id: string;
  type: 'plan' | 'plan_only';
  summary?: string;
  steps: PlanStep[];
  affectedSheets: string[];
  estimatedRows?: number;
  safestApproach?: string;
  /** The original user prompt, re-sent when running the plan as an action. */
  prompt: string;
  /** Tier 2 plan-only proposals (read-only — not queued for preview). */
  proposedActions?: SheetAction[];
  tier?: number;
}

export interface MatchResult {
  label: string;
  detail?: string;
  sheetName: string;
  row: number;
  col: number;
  colLetter?: string;
  rowNum?: number;
  /**
   * When set, the pointer selects from `col` to this 0-based column, so a
   * whole row of a table is shown and not one cell. Sent with answers that
   * were computed over a sheet's rows.
   */
  endCol?: number;
}

export interface MatchesBlock {
  id: string;
  type: 'matches';
  summary?: string;
  matches: MatchResult[];
}

/**
 * A completed GST reconciliation whose output sheet name collides with one
 * already in the workbook — rendered as Overwrite / Create-new buttons
 * (ActionResponseCard-style) rather than a typed chat question.
 */
export interface GstReconCollisionBlock {
  id: string;
  type: 'gst_recon_collision';
  collisionId: string;
  sheetName: string;
  /** Set while a button click's resolution is in flight — disables both buttons. */
  resolving?: boolean;
}

/** One books row that didn't cleanly match the portal — enough to identify and jump to it. */
export interface GstReconMissedRow {
  vendorName: string;
  gstin: string;
  reason: string;
  explanation: string;
  /** The books sheet's real name and Excel row number — always the CA's own register, never the GST portal sheet. */
  sheetName: string | null;
  row: number | null;
}

/**
 * The full list of books rows that didn't match the portal, shown alongside the
 * Accept/Reject action card so each one can be clicked to jump straight to that exact
 * cell in Purchase Register/Sales Register — never the GST portal sheet, since this
 * reconciliation is one-directional (books rows are what the CA owns and needs to verify).
 */
export interface GstReconMissedRowsBlock {
  id: string;
  type: 'gst_recon_missed_rows';
  rows: GstReconMissedRow[];
  /** Heading override (bank recon reuses this card for its unreconciled items). */
  title?: string;
}

export type TurnBlock =
  | StepBlock
  | ThinkingBlock
  | AnswerBlock
  | QuestionBlock
  | ActionBlock
  | StatusBlock
  | PlanBlock
  | MatchesBlock
  | GstReconCollisionBlock
  | GstReconMissedRowsBlock;

export type TurnPhase = 'processing' | 'awaiting_input' | 'complete' | 'error';

export interface ConversationTurn {
  id: string;
  userMessage: string;
  timestamp: Date;
  tabLabel: string;
  phase: TurnPhase;
  blocks: TurnBlock[];
  error?: string;
  /**
   * A ready-to-send follow-up that repairs cells the post-apply read-back found
   * holding Excel errors. Present only when the applied change actually left a
   * #REF!/#NAME?/... behind; the UI offers it rather than sending it, since the
   * write already landed and rewriting the user's cells is their call.
   * TASKS.md #168.
   */
  repairSuggestion?: RepairRequest;
  /**
   * Spilling formulas the read-back found blocked by content in their spill
   * area. Offered as a one-click clear that names the cells — TASKS.md #321.
   */
  spillBlockages?: SpillBlockage[];
  /**
   * An Accept that failed only because the step writes to sheets that don't
   * exist. Offered as one click that creates them and re-applies just what is
   * left of that step, instead of re-running the whole request. TASKS.md #346.
   */
  missingSheetRetry?: { blockId: string; sheets: string[] };
  /**
   * Files the user sent with this message (an attached bank statement). Shown
   * as chips in the message, like an attachment in any chat. A message with a
   * file cannot be edited or regenerated: the file itself is not kept.
   */
  attachmentNames?: string[];
  /**
   * Set when this turn carries on the turn before it instead of starting a new
   * exchange: the question typed with an attached statement is answered here,
   * after the import above it was accepted. It is drawn without a user message
   * of its own, so the two read as one exchange.
   */
  continuesTurnId?: string;
}

export function truncateTabLabel(text: string, max = 18): string {
  const trimmed = text.trim();
  if (trimmed.length <= max) return trimmed;
  return `${trimmed.slice(0, max - 1)}…`;
}

export function formatMessageTime(date: Date): string {
  return date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
}

/** Short relative label ("5d ago", "just now") for a response's own timestamp footer. */
export function formatRelativeTime(date: Date, now: Date = new Date()): string {
  const seconds = Math.max(0, Math.floor((now.getTime() - date.getTime()) / 1000));
  if (seconds < 45) return 'just now';
  const minutes = Math.floor(seconds / 60);
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  const days = Math.floor(hours / 24);
  if (days < 30) return `${days}d ago`;
  const months = Math.floor(days / 30);
  if (months < 12) return `${months}mo ago`;
  return `${Math.floor(months / 12)}y ago`;
}

/** Full date + time for the hover tooltip on a relative-time label. */
export function formatFullDateTime(date: Date): string {
  return date.toLocaleString([], {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  });
}
