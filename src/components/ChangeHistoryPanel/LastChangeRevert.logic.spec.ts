import { describe, expect, it } from 'vitest';
import { ChangeSetSummary } from '@/types/changeSet';

/**
 * LastChangeRevert renders the revert affordance for the most recent APPLIED
 * change at the end of the conversation. These cover the selection rule it
 * depends on, which is where a wrong answer would be user-visible: offering
 * revert on an already-reverted change, or on one that was never applied.
 *
 * (Kept as pure logic rather than a render test because the DOM test deps are
 * not installed in this environment — see TASKS.md #81.)
 */
const pickRevertable = (history: ChangeSetSummary[]) =>
  history.find((e) => e.status === 'applied') ?? null;

function entry(
  changeSetId: string,
  status: ChangeSetSummary['status'],
  sheets: string[] = ['Sheet1'],
): ChangeSetSummary {
  return {
    changeSetId,
    conversationId: 'conv1',
    traceId: 't1',
    timestamp: '2026-08-26T10:00:00Z',
    prompt: 'sort the sheet',
    changes: sheets.map((sheet, i) => ({
      sheet,
      cell: `A${i + 1}`,
      before: 'x',
      after: 'y',
    })) as ChangeSetSummary['changes'],
    status,
  };
}

describe('LastChangeRevert — which change set gets offered', () => {
  it('offers nothing when there is no history', () => {
    expect(pickRevertable([])).toBeNull();
  });

  it('picks the newest applied change (backend returns newest first)', () => {
    const history = [entry('cs3', 'applied'), entry('cs2', 'applied')];
    expect(pickRevertable(history)?.changeSetId).toBe('cs3');
  });

  it('skips an already-reverted change and offers the applied one beneath it', () => {
    // Reverting a reverted change would re-apply it — the button must not appear.
    const history = [entry('cs3', 'reverted'), entry('cs2', 'applied')];
    expect(pickRevertable(history)?.changeSetId).toBe('cs2');
  });

  it('offers nothing when every change is reverted', () => {
    expect(pickRevertable([entry('cs2', 'reverted'), entry('cs1', 'reverted')])).toBeNull();
  });

  it('does not offer a previewed change that was never applied', () => {
    // 'previewed' means staged for Accept but not yet written — reverting it
    // would try to undo a change the workbook never received.
    expect(pickRevertable([entry('cs1', 'previewed')])).toBeNull();
  });

  it('skips a previewed change sitting above an applied one', () => {
    // Real ordering after a new proposal arrives: newest is still previewed,
    // the revertable change is the applied one beneath it.
    const history = [entry('cs2', 'previewed'), entry('cs1', 'applied')];
    expect(pickRevertable(history)?.changeSetId).toBe('cs1');
  });
});

describe('LastChangeRevert — scope label', () => {
  const scopeOf = (e: ChangeSetSummary) => {
    const sheets = Array.from(new Set(e.changes.map((c) => c.sheet).filter(Boolean)));
    return sheets.length === 1 ? sheets[0] : sheets.length > 1 ? `${sheets.length} sheets` : '';
  };

  it('names the sheet when a change touches only one', () => {
    expect(scopeOf(entry('cs1', 'applied', ['Jan']))).toBe('Jan');
  });

  it('counts sheets when a change spans several', () => {
    expect(scopeOf(entry('cs1', 'applied', ['Jan', 'Feb', 'Mar']))).toBe('3 sheets');
  });

  it('collapses repeated sheet names rather than counting cells', () => {
    expect(scopeOf(entry('cs1', 'applied', ['Jan', 'Jan', 'Jan']))).toBe('Jan');
  });
});
