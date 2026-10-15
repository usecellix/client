// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { cleanup, render, screen } from '@testing-library/react';
import TurnRenderer from '@/components/ConversationPanel/TurnRenderer';
import { ActionBlock, AnswerBlock, ConversationTurn } from '@/types/conversationTurn';

const QUESTION = 'Which are the 10 largest debits?';

const answer = (turnId: string, content: string): AnswerBlock => ({
  id: `answer_${turnId}`,
  type: 'answer',
  content,
  revealState: 'complete',
});

const importCard = (overrides: Partial<ActionBlock> = {}): ActionBlock => ({
  id: 'actions_import_1',
  type: 'actions',
  actions: [{ type: 'CREATE_SHEET', sheetName: 'Bank Statement', name: 'Bank Statement' }] as ActionBlock['actions'],
  explanation: 'Create sheet "Bank Statement" with these transactions?',
  proposalStatus: 'pending',
  importSheetName: 'Bank Statement',
  ...overrides,
});

function turn(overrides: Partial<ConversationTurn>): ConversationTurn {
  return { id: 'turn_1', userMessage: QUESTION, timestamp: new Date(), tabLabel: 'x', phase: 'complete', blocks: [], ...overrides };
}

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
      onResolveGstReconCollision={vi.fn()}
      onRejectGstReconCollision={vi.fn()}
      onJumpToGstReconRow={vi.fn()}
    />,
  );
}

/**
 * A question sent with an attached statement is one exchange: one message with
 * the file on it, the import card, then the answer. ATTACHMENT_EXTRACTION_PLAN.md §6.5.
 */
describe('TurnRenderer — a message sent with a file', () => {
  afterEach(() => cleanup());

  const importTurn = (card: ActionBlock) =>
    turn({
      attachmentNames: ['statement.pdf'],
      blocks: [answer('turn_1', 'I read statement.pdf and found 834 transactions.'), card],
    });

  it('shows the question as the message, with the file attached to it', () => {
    const { container } = renderTurn(importTurn(importCard({ followUp: { prompt: QUESTION, mode: 'ask' } })));

    const message = container.querySelector('.cellix-user-msg') as HTMLElement;
    expect(message.querySelector('.cellix-user-msg-text')?.textContent).toBe(QUESTION);
    expect(message.querySelector('.cellix-user-msg-file')?.textContent).toBe('statement.pdf');
    expect(message.querySelector('.cellix-user-msg-file .is-pdf')).not.toBeNull();
  });

  it('offers no edit or regenerate on a message that carried a file', () => {
    renderTurn(importTurn(importCard()));
    expect(screen.queryByRole('button', { name: 'Message actions' })).toBeNull();
  });

  it('holds back the footer and suggestions while the question is still to be answered', () => {
    const { container } = renderTurn(importTurn(importCard({ followUp: { prompt: QUESTION, mode: 'ask' } })));
    expect(container.querySelector('.cellix-response-footer, .cellix-followups, [class*="follow-up"]')).toBeNull();
    expect(screen.queryByText(/just now/i)).toBeNull();
  });

  it('shows the footer on an import that came with no question', () => {
    renderTurn(importTurn(importCard()));
    expect(screen.queryByText(/just now/i)).not.toBeNull();
  });

  it('draws the answer turn with no message of its own, joined to the import above', () => {
    const { container } = renderTurn(
      turn({ id: 'turn_2', continuesTurnId: 'turn_1', blocks: [answer('turn_2', 'The largest debit is 30,000.')] }),
    );

    expect(container.querySelector('.cellix-turn')?.className).toContain('is-continuation');
    expect(container.querySelector('.cellix-user-msg')).toBeNull();
    expect(screen.queryByText(QUESTION)).toBeNull();
    expect(screen.getByText(/The largest debit is 30,000\./)).toBeTruthy();
  });

  it('leaves an ordinary message exactly as it was', () => {
    const { container } = renderTurn(turn({ blocks: [answer('turn_1', 'Done.')] }));
    expect(container.querySelector('.cellix-turn')?.className).toBe('cellix-turn');
    expect(container.querySelector('.cellix-user-msg-file')).toBeNull();
    expect(screen.getByRole('button', { name: 'Message actions' })).toBeTruthy();
  });
});
