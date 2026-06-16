import { ConversationTurn } from '@/types/conversationTurn';

/** True when thinking, steps, and answer typing have finished for a turn. */
export function isTurnPresentationComplete(turn: ConversationTurn): boolean {
  if (turn.phase === 'processing' || turn.phase === 'error') {
    return false;
  }

  for (const block of turn.blocks) {
    if (block.type === 'thinking' && block.loading) {
      return false;
    }
    if (block.type === 'answer' && block.revealState !== 'complete') {
      return false;
    }
    if (block.type === 'step' && (block.phase === 'running' || block.phase === 'revealed')) {
      return false;
    }
    if (block.type === 'status' && block.pulsing) {
      return false;
    }
  }

  return true;
}
