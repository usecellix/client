import { describe, expect, it } from 'vitest';
import { isLikelyChitchat } from './statusMessage';

/**
 * User report: sending "hi" showed the full "Reading your worksheet…
 * Analyzing your spreadsheet structure and understanding the data…"
 * choreography (with several seconds of hard-minimum artificial delay) even
 * though the backend's own CHITCHAT route (see `chitchat-prompt.ts`) skips
 * workbook context entirely and replies in well under a second. This mirrors
 * that backend classification client-side, conservatively, so the frontend
 * can skip the same wasted wait.
 */
describe('isLikelyChitchat', () => {
  it('recognizes bare greetings', () => {
    expect(isLikelyChitchat('hi')).toBe(true);
    expect(isLikelyChitchat('Hello')).toBe(true);
    expect(isLikelyChitchat('hey there')).toBe(true);
    expect(isLikelyChitchat('  hi  ')).toBe(true);
  });

  it('recognizes thanks/goodbyes/identity questions', () => {
    expect(isLikelyChitchat('thanks')).toBe(true);
    expect(isLikelyChitchat('thank you!')).toBe(true);
    expect(isLikelyChitchat('bye')).toBe(true);
    expect(isLikelyChitchat('who are you?')).toBe(true);
    expect(isLikelyChitchat('what can you do')).toBe(true);
  });

  it('recognizes short acknowledgements and emoji-only messages', () => {
    expect(isLikelyChitchat('ok')).toBe(true);
    expect(isLikelyChitchat('cool!')).toBe(true);
    expect(isLikelyChitchat('👍')).toBe(true);
  });

  it('does NOT match anything referencing the spreadsheet or asking for work', () => {
    expect(isLikelyChitchat('create a table with headers')).toBe(false);
    expect(isLikelyChitchat('sum column B')).toBe(false);
    expect(isLikelyChitchat('what is the total in this sheet')).toBe(false);
    expect(isLikelyChitchat('explain this data')).toBe(false);
    expect(isLikelyChitchat('fix the formula in A1')).toBe(false);
  });

  it('does not falsely match a greeting-prefixed real request (a real task, not chitchat)', () => {
    expect(isLikelyChitchat('hi, can you build a dashboard for me')).toBe(false);
    expect(isLikelyChitchat('hello, please sort column A')).toBe(false);
  });

  it('rejects long messages even if superficially greeting-shaped', () => {
    const long = 'hi ' + 'x'.repeat(50);
    expect(isLikelyChitchat(long)).toBe(false);
  });

  it('rejects empty/whitespace-only input', () => {
    expect(isLikelyChitchat('')).toBe(false);
    expect(isLikelyChitchat('   ')).toBe(false);
  });
});
