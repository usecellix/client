// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import TurnRenderer from '@/components/ConversationPanel/TurnRenderer';
import { ActionBlock, ConversationTurn } from '@/types/conversationTurn';

/**
 * TASKS.md #404 — live: after a successful Revert the step card went on reading
 * "Applied", and reopening the message menu offered "Revert this change" again.
 * The revert control kept "reverted" in its own state, which is lost every time
 * the menu closes. It is now a field on the block, and both the card and the
 * menu read it from there.
 */
const step = (id: string, overrides: Partial<ActionBlock> = {}): ActionBlock => ({
  id: `actions_${id}`,
  type: 'actions',
  actions: [{ type: 'FORMAT_RANGE', sheetName: 'S', row: 0, col: 0, format: { bold: true } }] as ActionBlock['actions'],
  explanation: 'Format the header',
  proposalStatus: 'accepted',
  changeSetId: `cs_${id}`,
  ...overrides,
});

const turn = (blocks: ActionBlock[]): ConversationTurn => ({
  id: 'turn_1',
  userMessage: 'Make the header bold',
  timestamp: new Date(),
  tabLabel: 'x',
  phase: 'complete',
  blocks,
});

function renderTurn(t: ConversationTurn) {
  return render(
    <TurnRenderer
      turn={t}
      isActive
      isWaiting={false}
      showActionButtons
      onAcceptActions={vi.fn()}
      onRejectActions={vi.fn()}
      onAnswerQuestion={vi.fn()}
      onToggleThinking={vi.fn()}
      onAnswerComplete={vi.fn()}
      onFollowUp={vi.fn()}
      onRegenerate={vi.fn()}
      onRunAsAction={vi.fn()}
      onRevertChangeSet={vi.fn().mockResolvedValue(undefined)}
      onResolveGstReconCollision={vi.fn()}
      onJumpToGstReconRow={vi.fn()}
    />,
  );
}

const toggleMenu = () => fireEvent.click(screen.getByRole('button', { name: 'Message actions' }));

describe('TurnRenderer — a change that was reverted', () => {
  afterEach(() => cleanup());

  it('offers Revert for an applied change, and its card says Applied', () => {
    const { container } = renderTurn(turn([step('1')]));
    expect(container.querySelector('.cellix-action-status.is-applied')?.textContent).toBe('Applied');

    toggleMenu();
    expect(screen.getByRole('menuitem', { name: /revert this change/i })).toBeTruthy();
  });

  it('says Reverted on the card and in the menu, and offers no Revert', () => {
    const { container } = renderTurn(turn([step('1', { reverted: true })]));

    expect(container.querySelector('.cellix-action-status.is-reverted')?.textContent).toBe('Reverted');
    expect(container.querySelector('.cellix-action-status.is-applied')).toBeNull();

    toggleMenu();
    expect(screen.queryByRole('menuitem', { name: /revert this change/i })).toBeNull();
    expect(screen.getByRole('menu').textContent).toContain('Reverted');
  });

  it('still says Reverted when the menu is closed and opened again', () => {
    renderTurn(turn([step('1', { reverted: true })]));

    toggleMenu();
    toggleMenu();
    expect(screen.queryByRole('menu')).toBeNull();
    toggleMenu();

    expect(screen.getByRole('menu').textContent).toContain('Reverted');
    expect(screen.queryByRole('menuitem', { name: /revert this change/i })).toBeNull();
  });

  it('keeps offering Revert for the steps of a message that are not yet reverted', () => {
    const { container } = renderTurn(turn([step('1'), step('2', { reverted: true })]));

    expect(container.querySelectorAll('.cellix-action-status.is-applied')).toHaveLength(1);
    expect(container.querySelectorAll('.cellix-action-status.is-reverted')).toHaveLength(1);

    toggleMenu();
    expect(screen.getByRole('menuitem', { name: /revert this change/i })).toBeTruthy();
  });
});
