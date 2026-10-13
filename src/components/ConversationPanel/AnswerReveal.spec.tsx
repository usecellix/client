// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, cleanup, render, screen } from '@testing-library/react';
import AnswerReveal from '@/components/ConversationPanel/AnswerReveal';
import { MatchResult } from '@/types/conversationTurn';

vi.mock('@/services/rangeFetchService', () => ({ navigateToCell: vi.fn(async () => undefined) }));

const LINE_1 = '1. 24-Jul-2026 · Debit 30,000.00 (row 445)';
const LINE_2 = '2. 14-Apr-2026 · Debit 20,000.00 (row 16)';
const ANSWER = `**2 largest by Debit**\n${LINE_1}\n${LINE_2}\n\nWorked out from all 834 rows of Bank Statement.`;

const pointer = (rowNum: number): MatchResult => ({ label: `row ${rowNum}`, sheetName: 'Bank Statement', row: rowNum - 1, col: 0, rowNum, endCol: 8 });
const POINTERS = [pointer(445), pointer(16)];
const TICK = 24;

/** Advances the typing animation to just after `text` has been typed. */
function typeUntil(text: string) {
  const upTo = ANSWER.indexOf(text) + text.length;
  act(() => {
    vi.advanceTimersByTime(TICK * upTo);
  });
}

/**
 * Row pointers appear while the answer is being typed, and the finished answer
 * is the same block that was typed, not a second one drawn over it. TASKS.md #392.
 */
describe('AnswerReveal — typing an answer that names its rows', () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => {
    cleanup();
    vi.useRealTimers();
  });

  const buttons = () => screen.queryAllByRole('button').map((b) => b.textContent);

  it('shows each pointer as soon as its line has been typed', () => {
    const { container } = render(<AnswerReveal content={ANSWER} revealState="typing" matches={POINTERS} />);

    typeUntil('**2 largest by Debit**');
    expect(buttons()).toEqual([]);

    typeUntil(LINE_1);
    expect(buttons()).toEqual(['row 445']);
    expect(container.textContent).not.toContain('(row 445)');

    typeUntil(LINE_2);
    expect(buttons()).toEqual(['row 445', 'row 16']);
  });

  it('never shows a half-typed "(row" before the pointer', () => {
    const { container } = render(<AnswerReveal content={ANSWER} revealState="typing" matches={POINTERS} />);
    const start = ANSWER.indexOf('(row 445)');

    // Step through the marker one character at a time.
    act(() => {
      vi.advanceTimersByTime(TICK * start);
    });
    for (let typed = 1; typed < '(row 445)'.length; typed++) {
      act(() => {
        vi.advanceTimersByTime(TICK);
      });
      expect(container.textContent).not.toMatch(/\(r?o?w?\s?\d*$/);
      expect(buttons()).toEqual([]);
    }
    act(() => {
      vi.advanceTimersByTime(TICK);
    });
    expect(buttons()).toEqual(['row 445']);
  });

  it('keeps the same block when typing finishes, with nothing drawn a second time', () => {
    const onComplete = vi.fn();
    const { container, rerender } = render(
      <AnswerReveal content={ANSWER} revealState="typing" matches={POINTERS} onComplete={onComplete} />,
    );
    const whileTyping = container.querySelector('.cellix-response-output');
    typeUntil(LINE_1);
    const firstPointer = screen.getByRole('button', { name: 'Show row 445 in Bank Statement' });

    act(() => {
      vi.advanceTimersByTime(TICK * ANSWER.length);
    });
    expect(onComplete).toHaveBeenCalledTimes(1);
    rerender(<AnswerReveal content={ANSWER} revealState="complete" matches={POINTERS} onComplete={onComplete} />);

    // Same DOM nodes as during typing: the block was not replaced.
    expect(container.querySelector('.cellix-response-output')).toBe(whileTyping);
    expect(screen.getByRole('button', { name: 'Show row 445 in Bank Statement' })).toBe(firstPointer);
    expect(container.querySelectorAll('.cellix-response-output')).toHaveLength(1);
    expect(buttons()).toEqual(['row 445', 'row 16']);
  });

  it('carries on typing when the thread re-renders, and does not start again from the first character', () => {
    const text = 'The sheet has 834 rows.';
    const first = vi.fn();
    const second = vi.fn();
    const { container, rerender } = render(<AnswerReveal content={text} revealState="typing" onComplete={first} />);

    act(() => {
      vi.advanceTimersByTime(TICK * 12);
    });
    expect(container.textContent).toBe('The sheet ha');

    // What a credits update or the end-of-stream event does: same answer, new callback.
    rerender(<AnswerReveal content={text} revealState="typing" onComplete={second} />);
    expect(container.textContent).toBe('The sheet ha');

    act(() => {
      vi.advanceTimersByTime(TICK * 3);
    });
    expect(container.textContent).toBe('The sheet has 8');

    act(() => {
      vi.advanceTimersByTime(TICK * text.length);
    });
    // Finished once, reported to the callback that is current.
    expect(second).toHaveBeenCalledTimes(1);
    expect(first).not.toHaveBeenCalled();
  });

  it('types an ordinary answer as before and keeps its block too', () => {
    const text = 'The sheet has 834 rows.';
    const { container, rerender } = render(<AnswerReveal content={text} revealState="typing" />);
    const whileTyping = container.querySelector('.cellix-response-output');

    act(() => {
      vi.advanceTimersByTime(TICK * 8);
    });
    expect(container.textContent).toBe('The shee');
    expect(container.querySelector('.cellix-answer-streaming')).not.toBeNull();

    act(() => {
      vi.advanceTimersByTime(TICK * text.length);
    });
    rerender(<AnswerReveal content={text} revealState="complete" />);
    expect(container.textContent).toContain(text);
    expect(container.querySelector('.cellix-response-output')).toBe(whileTyping);
    expect(container.querySelector('.cellix-answer-streaming')).toBeNull();
  });

  it('still adds a find result\'s cell list only once its sentence is complete', () => {
    const found: MatchResult = { label: '593111', sheetName: 'Bank Statement', row: 4, col: 2, colLetter: 'C', rowNum: 5 };
    const text = 'Found 1 match for 593111.';
    const { container, rerender } = render(<AnswerReveal content={text} revealState="typing" matches={[found]} />);

    act(() => {
      vi.advanceTimersByTime(TICK * 10);
    });
    expect(container.querySelector('.cellix-find-pointer')).toBeNull();

    rerender(<AnswerReveal content={text} revealState="complete" matches={[found]} />);
    expect(container.querySelector('.cellix-find-pointer')?.textContent).toBe('Bank Statement!C5');
  });

  it('shows nothing while hidden', () => {
    const { container } = render(<AnswerReveal content={ANSWER} revealState="hidden" matches={POINTERS} />);
    expect(container.innerHTML).toBe('');
  });
});
