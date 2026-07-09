import { describe, expect, it } from 'vitest';
import { ConversationTurn } from '@/types/conversationTurn';
import { isTurnPresentationComplete } from './turnPresentation';

describe('isTurnPresentationComplete', () => {
  const baseTurn = (overrides: Partial<ConversationTurn> = {}): ConversationTurn => ({
    id: 'turn_1',
    userMessage: 'test',
    timestamp: new Date(),
    tabLabel: 'test',
    phase: 'complete',
    blocks: [],
    ...overrides,
  });

  it('returns false while turn is processing', () => {
    expect(isTurnPresentationComplete(baseTurn({ phase: 'processing' }))).toBe(false);
  });

  it('returns false while thinking is loading', () => {
    expect(
      isTurnPresentationComplete(
        baseTurn({
          blocks: [{ id: 't1', type: 'thinking', content: 'Analyzing…', expanded: true, loading: true }],
        }),
      ),
    ).toBe(false);
  });

  it('returns false while answer is typing', () => {
    expect(
      isTurnPresentationComplete(
        baseTurn({
          blocks: [
            { id: 'a1', type: 'answer', content: 'Done', revealState: 'typing' },
          ],
        }),
      ),
    ).toBe(false);
  });

  it('returns true when thinking finished and answer complete', () => {
    expect(
      isTurnPresentationComplete(
        baseTurn({
          blocks: [
            { id: 't1', type: 'thinking', content: 'Done', expanded: false, loading: false },
            { id: 'a1', type: 'answer', content: 'Done', revealState: 'complete' },
          ],
        }),
      ),
    ).toBe(true);
  });

  it('returns false while a step is running', () => {
    expect(
      isTurnPresentationComplete(
        baseTurn({
          blocks: [{ id: 's1', type: 'step', label: 'Reading', phase: 'running' }],
        }),
      ),
    ).toBe(false);
  });

  it('returns false while status is pulsing', () => {
    expect(
      isTurnPresentationComplete(
        baseTurn({
          blocks: [{ id: 'st1', type: 'status', label: 'Working…', pulsing: true }],
        }),
      ),
    ).toBe(false);
  });
});
