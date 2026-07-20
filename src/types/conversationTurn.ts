import { SheetAction } from '@/hooks/useSseStream';
import { CellChange } from '@/types/changeSet';

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
}

export function truncateTabLabel(text: string, max = 18): string {
  const trimmed = text.trim();
  if (trimmed.length <= max) return trimmed;
  return `${trimmed.slice(0, max - 1)}…`;
}

export function formatMessageTime(date: Date): string {
  return date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
}
