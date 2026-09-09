import type { RepairRequest } from '@/services/repairRequest';
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
}

export interface MatchesBlock {
  id: string;
  type: 'matches';
  summary?: string;
  matches: MatchResult[];
}

export type TurnBlock =
  | StepBlock
  | ThinkingBlock
  | AnswerBlock
  | QuestionBlock
  | ActionBlock
  | StatusBlock
  | PlanBlock
  | MatchesBlock;

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
