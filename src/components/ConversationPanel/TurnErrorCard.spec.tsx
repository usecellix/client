// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { TurnErrorCard } from './TurnErrorCard';

/** TASKS.md #321 — the red card offers the fix, not just the report. */
describe('TurnErrorCard', () => {
  afterEach(() => cleanup());

  const blockage = { sheet: 'Main', anchor: 'A19', blockers: ['G19', 'I19', 'M19'] };

  it('offers to clear exactly the blocking cells, and confirms once the list fills in', async () => {
    const clear = vi.fn().mockResolvedValue(true);
    render(
      <TurnErrorCard
        message="Applied, but the list at Main!A19 can't fill in"
        spillBlockages={[blockage]}
        clearSpillBlockers={clear}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: 'Clear G19, I19 and M19' }));

    await waitFor(() => expect(screen.getByRole('status').textContent).toMatch(/now fills in/));
    expect(clear).toHaveBeenCalledWith(blockage);
  });

  it('says so when clearing did not unblock it, instead of claiming success', async () => {
    render(
      <TurnErrorCard
        message="x"
        spillBlockages={[blockage]}
        clearSpillBlockers={vi.fn().mockResolvedValue(false)}
      />,
    );
    fireEvent.click(screen.getByRole('button', { name: /Clear/ }));
    await waitFor(() => expect(screen.getByText(/still can't fill in/)).toBeTruthy());
    expect(screen.queryByRole('status')).toBeNull();
  });

  it('offers the formula repair (never rendered before) when there is no spill fix', () => {
    const onRepair = vi.fn();
    render(
      <TurnErrorCard
        message="Applied, but 1 formula cell(s) returned an error: Main!B2 #REF!."
        repairSuggestion={{ prompt: 'fix B2', cellCount: 1, errors: ['#REF!'] }}
        onRepair={onRepair}
      />,
    );
    fireEvent.click(screen.getByRole('button', { name: 'Fix this formula' }));
    expect(onRepair).toHaveBeenCalledWith('fix B2');
  });

  it('is plain text when there is nothing to offer', () => {
    render(<TurnErrorCard message="Something went wrong." />);
    expect(screen.getByRole('alert').textContent).toBe('Something went wrong.');
    expect(screen.queryByRole('button')).toBeNull();
  });
});
