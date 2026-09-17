// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { TurnRevertControl } from '@/components/ConversationPanel/TurnRenderer';
import * as auditService from '@/services/auditService';

/**
 * TASKS.md #242 — a live run showed a revert control rendering for an
 * AUTO_FILTER change set (correctly flagged `irreversibleActionTypes:
 * ["AUTO_FILTER"]` server-side, confirmed by reading the change_sets document
 * directly) that then failed with HTTP 422 — the server's own fail-closed
 * self-verification (change-set.controller.ts's RevertVerificationError
 * handling) correctly refusing a revert it cannot perform correctly. No data
 * was ever at risk, but the button relabeled itself "Retry revert", inviting
 * a click that is guaranteed to fail the exact same way every time.
 *
 * This does not fix why the control rendered in the first place (this
 * session's static tracing of the SSE payload -> pending block ->
 * createActionBlock -> accept-transition chain found every step correctly
 * threading `irreversibleActionTypes` through, so the exact trigger is still
 * open — see the task row) — it fixes the dead-end UX so a 422 is never
 * presented as something worth retrying.
 */
describe('TurnRevertControl (#242)', () => {
  afterEach(() => {
    cleanup();
    vi.restoreAllMocks();
  });

  it('shows a terminal "Can\'t be undone" state on a 422, not "Retry revert"', async () => {
    vi.spyOn(auditService, 'revertChangeSet').mockRejectedValue(
      new Error('Audit API 422: {"message":"No inverse exists","code":"NOT_REVERSIBLE"}'),
    );
    const onRevert = vi.fn();

    render(<TurnRevertControl changeSetIds={['cs_1']} onRevert={onRevert} />);
    fireEvent.click(screen.getByRole('menuitem', { name: /revert this change/i }));

    await waitFor(() => expect(screen.getByText(/can't be undone/i)).toBeTruthy());
    expect(screen.queryByText(/retry revert/i)).toBeNull();
    // Terminal — no button to click again.
    expect(screen.queryByRole('menuitem')).toBeNull();
    expect(onRevert).not.toHaveBeenCalled();
  });

  it('still offers "Retry revert" for a non-422 (potentially transient) failure', async () => {
    vi.spyOn(auditService, 'revertChangeSet').mockRejectedValue(new Error('Audit API 500: timeout'));
    render(<TurnRevertControl changeSetIds={['cs_1']} onRevert={vi.fn()} />);

    fireEvent.click(screen.getByRole('menuitem', { name: /revert this change/i }));

    await waitFor(() => expect(screen.getByText(/retry revert/i)).toBeTruthy());
    expect(screen.queryByText(/can't be undone/i)).toBeNull();
  });

  it('shows Reverted and calls onRevert on success', async () => {
    const inverseActions = [{ type: 'AUTO_FILTER', sheetName: 'S', range: 'A1:A1' }] as never;
    vi.spyOn(auditService, 'revertChangeSet').mockResolvedValue({
      changeSet: {} as never,
      inverseActions,
    });
    const onRevert = vi.fn().mockResolvedValue(undefined);

    render(<TurnRevertControl changeSetIds={['cs_1']} onRevert={onRevert} />);
    fireEvent.click(screen.getByRole('menuitem', { name: /revert this change/i }));

    await waitFor(() => expect(screen.getByText(/^reverted$/i)).toBeTruthy());
    expect(onRevert).toHaveBeenCalledWith('cs_1', inverseActions);
  });
});

/**
 * A single user prompt ("Copy the Purchase Register sheet and name it March
 * Copy") can produce several stepwise change sets — one for creating the
 * sheet, one for writing its content — each its own ActionBlock. A live run
 * showed two separate "Revert this change" items in the same menu for what
 * the user experiences as one action, and asked for it to be treated as one.
 */
describe('TurnRevertControl — multiple change sets treated as one action', () => {
  afterEach(() => {
    cleanup();
    vi.restoreAllMocks();
  });

  it('reverts multiple change sets in reverse (last-applied-first) order as a single click', async () => {
    const calls: string[] = [];
    vi.spyOn(auditService, 'revertChangeSet').mockImplementation(async (id: string) => {
      calls.push(id);
      return { changeSet: {} as never, inverseActions: [{ type: 'SET_CELL', id } as never] };
    });
    const onRevert = vi.fn().mockResolvedValue(undefined);

    render(<TurnRevertControl changeSetIds={['cs_step1', 'cs_step2']} onRevert={onRevert} />);
    fireEvent.click(screen.getByRole('menuitem', { name: /revert this change/i }));

    await waitFor(() => expect(screen.getByText(/^reverted$/i)).toBeTruthy());
    // step2 (written after step1) undoes first, step1 (the sheet create) last.
    expect(calls).toEqual(['cs_step2', 'cs_step1']);
    expect(onRevert).toHaveBeenNthCalledWith(1, 'cs_step2', [{ type: 'SET_CELL', id: 'cs_step2' }]);
    expect(onRevert).toHaveBeenNthCalledWith(2, 'cs_step1', [{ type: 'SET_CELL', id: 'cs_step1' }]);
  });

  it('stops at the first failure and reports it, without reverting earlier steps after it', async () => {
    const calls: string[] = [];
    vi.spyOn(auditService, 'revertChangeSet').mockImplementation(async (id: string) => {
      calls.push(id);
      if (id === 'cs_step1') {
        throw new Error('Audit API 422: {"message":"No inverse exists"}');
      }
      return { changeSet: {} as never, inverseActions: [] };
    });
    const onRevert = vi.fn().mockResolvedValue(undefined);

    render(<TurnRevertControl changeSetIds={['cs_step1', 'cs_step2']} onRevert={onRevert} />);
    fireEvent.click(screen.getByRole('menuitem', { name: /revert this change/i }));

    await waitFor(() => expect(screen.getByText(/can't be undone/i)).toBeTruthy());
    // step2 reverted successfully before step1 failed.
    expect(calls).toEqual(['cs_step2', 'cs_step1']);
    expect(onRevert).toHaveBeenCalledTimes(1);
    expect(onRevert).toHaveBeenCalledWith('cs_step2', []);
  });
});
