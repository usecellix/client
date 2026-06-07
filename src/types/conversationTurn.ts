import { SheetAction } from '@/hooks/useSseStream';

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
}

export interface StatusBlock {
  id: string;
  type: 'status';
  label: string;
  pulsing: boolean;
  visible?: boolean;
}

export type TurnBlock =
  | StepBlock
  | ThinkingBlock
  | AnswerBlock
  | QuestionBlock
  | ActionBlock
  | StatusBlock;

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
