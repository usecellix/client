// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import ResponseOutput from '@/components/ConversationPanel/ResponseOutput';
import { navigateToCell } from '@/services/rangeFetchService';
import { MatchResult } from '@/types/conversationTurn';

vi.mock('@/services/rangeFetchService', () => ({ navigateToCell: vi.fn(async () => undefined) }));
const navigate = vi.mocked(navigateToCell);

const ANSWER = [
  '**3 largest by Debit**',
  '1. 24-Jul-2026 · UPIOUT/1268 · Debit 30,000.00 (row 445)',
  '2. 14-Apr-2026 · UPIOUT/1215 · Debit 20,000.00 (row 16)',
  '3. 19-May-2026 · UPIOUT/1233 · Debit 20,000.00 (row 141)',
  '2 rows more have the same Debit (20,000.00) as the last one listed.',
  '',
  'Worked out from all 834 rows of Bank Statement.',
].join('\n');

const pointer = (rowNum: number): MatchResult => ({
  label: `row ${rowNum}`,
  sheetName: 'Bank Statement',
  row: rowNum - 1,
  col: 0,
  colLetter: 'A',
  rowNum,
  endCol: 8,
});
const POINTERS = [pointer(445), pointer(16), pointer(141)];

/** Each row named in a computed answer is a pointer that shows that row in the workbook. TASKS.md #391. */
describe('ResponseOutput — rows named in a computed answer', () => {
  beforeEach(() => navigate.mockClear());
  afterEach(() => cleanup());

  it('turns every "(row N)" into a pointer, in place', () => {
    const { container } = render(<ResponseOutput content={ANSWER} matches={POINTERS} />);

    expect(screen.getAllByRole('button').map((b) => b.textContent)).toEqual(['row 445', 'row 16', 'row 141']);
    // The answer keeps its lines and its heading; only the row numbers changed.
    expect(container.querySelector('strong')?.textContent).toBe('3 largest by Debit');
    expect(container.textContent).toContain('1. 24-Jul-2026 · UPIOUT/1268 · Debit 30,000.00 row 445');
    expect(container.textContent).toContain('Worked out from all 834 rows of Bank Statement.');
    expect(container.textContent).not.toContain('(row 445)');
    // Not the one-line "find" layout, which lists cell references after the text.
    expect(container.querySelector('.cellix-find-pointer')).toBeNull();
  });

  it('selects the whole row on the right sheet when a pointer is clicked', () => {
    render(<ResponseOutput content={ANSWER} matches={POINTERS} />);

    fireEvent.click(screen.getByRole('button', { name: 'Show row 16 in Bank Statement' }));

    expect(navigate).toHaveBeenCalledTimes(1);
    // Sheet row 16 is index 15; columns A to I.
    expect(navigate).toHaveBeenCalledWith('Bank Statement', 15, 0, 8);
  });

  it('leaves a row number as text when no pointer came for it', () => {
    const { container } = render(<ResponseOutput content={ANSWER} matches={[pointer(445)]} />);

    expect(screen.getAllByRole('button').map((b) => b.textContent)).toEqual(['row 445']);
    expect(container.textContent).toContain('Debit 20,000.00 (row 16)');
  });

  it('points at the same row from each place the answer names it', () => {
    const twice = '**Lowest Balance: 2,121.82**\n   18-May-2026 (row 9).\n\n**1 largest by Debit**\n1. 18-May-2026 · Debit 20,000.00 (row 9)';
    render(<ResponseOutput content={twice} matches={[pointer(9)]} />);
    expect(screen.getAllByRole('button', { name: 'Show row 9 in Bank Statement' })).toHaveLength(2);
  });

  it('does not throw when the sheet is gone', async () => {
    navigate.mockRejectedValueOnce(new Error('ItemNotFound'));
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    render(<ResponseOutput content={ANSWER} matches={POINTERS} />);

    fireEvent.click(screen.getByRole('button', { name: 'Show row 445 in Bank Statement' }));
    await Promise.resolve();

    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
  });

  it('keeps the existing one-line layout for a find result', () => {
    const found: MatchResult = { label: '593111', sheetName: 'Bank Statement', row: 4, col: 2, colLetter: 'C', rowNum: 5 };
    const { container } = render(<ResponseOutput content="Found 1 match for 593111." matches={[found]} />);

    expect(container.querySelector('.cellix-find-pointer')?.textContent).toBe('Bank Statement!C5');
    expect(container.querySelector('.cellix-row-pointer')).toBeNull();
    fireEvent.click(container.querySelector('.cellix-find-pointer') as HTMLElement);
    expect(navigate).toHaveBeenCalledWith('Bank Statement', 4, 2);
  });

  it('renders an answer with no pointers exactly as before', () => {
    const { container } = render(<ResponseOutput content={ANSWER} />);
    expect(screen.queryAllByRole('button')).toHaveLength(0);
    expect(container.textContent).toContain('(row 445)');
  });
});
