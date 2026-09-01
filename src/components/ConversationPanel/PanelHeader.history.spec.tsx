// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';

/**
 * TASKS.md #172 — the chat-history panel's list-rendering and click-to-reopen
 * logic, which is the part of the feature unit tests can actually own. The
 * end-to-end path (log in → converse → reload → reopen) is manual per this
 * repo's UI-change convention.
 *
 * `useSession` and `signOutUser` are mocked because this suite is about the
 * history menu, not auth; `CheckpointPanel` is never rendered here
 * (`showCheckpointsButton` defaults false).
 */
vi.mock('@/auth/auth-client', () => ({
  useSession: () => ({ data: { user: { email: 'user@example.com' } } }),
}));
vi.mock('@/auth/useAuth', () => ({ signOutUser: vi.fn() }));
vi.mock('@/services/conversationHistoryService', () => ({
  fetchConversationHistory: vi.fn(),
}));

import PanelHeader from '@/components/ConversationPanel/PanelHeader';
import { fetchConversationHistory } from '@/services/conversationHistoryService';
import { ChatSession } from '@/types/chatSession';

const fetchHistoryMock = vi.mocked(fetchConversationHistory);

function historyEntry(overrides: Record<string, unknown> = {}) {
  return {
    conversationId: 'conv_past',
    title: 'Add a Remarks column',
    firstMessage: 'Add a Remarks column',
    lastMessage: 'Added.',
    messageCount: 4,
    status: 'completed',
    updatedAt: new Date().toISOString(),
    ...overrides,
  };
}

function renderHeader(
  overrides: {
    sessions?: ChatSession[];
    activeSessionId?: string | null;
    onOpenHistoryConversation?: (id: string) => Promise<boolean>;
  } = {},
) {
  const onOpenHistoryConversation =
    overrides.onOpenHistoryConversation ?? vi.fn().mockResolvedValue(true);

  render(
    <PanelHeader
      sessions={overrides.sessions ?? []}
      activeSessionId={overrides.activeSessionId ?? null}
      isWaitingForResponse={false}
      onSelectSession={vi.fn()}
      onCloseSession={vi.fn()}
      onNewChat={vi.fn()}
      onOpenHistoryConversation={onOpenHistoryConversation}
      conversationId={null}
      onRestoreCheckpoint={vi.fn()}
    />,
  );

  return { onOpenHistoryConversation };
}

function openHistoryMenu() {
  fireEvent.click(screen.getByLabelText('Chat history'));
}

beforeEach(() => {
  fetchHistoryMock.mockReset();
  fetchHistoryMock.mockResolvedValue({ conversations: [], nextCursor: null });
});

afterEach(cleanup);

describe('PanelHeader chat history (TASKS.md #172)', () => {
  it('does not fetch history until the menu is opened', () => {
    renderHeader();
    expect(fetchHistoryMock).not.toHaveBeenCalled();
  });

  it('fetches and lists past conversations when opened', async () => {
    fetchHistoryMock.mockResolvedValue({
      conversations: [historyEntry()],
      nextCursor: null,
    });

    renderHeader();
    openHistoryMenu();

    expect(await screen.findByText('Add a Remarks column')).toBeTruthy();
    expect(fetchHistoryMock).toHaveBeenCalledTimes(1);
  });

  it('reopens a conversation by conversationId when a row is clicked', async () => {
    fetchHistoryMock.mockResolvedValue({
      conversations: [historyEntry()],
      nextCursor: null,
    });
    const { onOpenHistoryConversation } = renderHeader();
    openHistoryMenu();

    fireEvent.click(await screen.findByText('Add a Remarks column'));

    await waitFor(() => {
      expect(onOpenHistoryConversation).toHaveBeenCalledWith('conv_past');
    });
  });

  it('closes the menu after a successful reopen', async () => {
    fetchHistoryMock.mockResolvedValue({
      conversations: [historyEntry()],
      nextCursor: null,
    });
    renderHeader();
    openHistoryMenu();

    fireEvent.click(await screen.findByText('Add a Remarks column'));

    await waitFor(() => {
      expect(screen.queryByLabelText('Search chats')).toBeNull();
    });
  });

  it('keeps the menu open and reports failure when reopening fails', async () => {
    fetchHistoryMock.mockResolvedValue({
      conversations: [historyEntry()],
      nextCursor: null,
    });
    renderHeader({ onOpenHistoryConversation: vi.fn().mockResolvedValue(false) });
    openHistoryMenu();

    fireEvent.click(await screen.findByText('Add a Remarks column'));

    expect(await screen.findByText("Couldn't open that chat.")).toBeTruthy();
    expect(screen.queryByLabelText('Search chats')).toBeTruthy();
  });

  it('shows an empty state rather than a blank menu', async () => {
    renderHeader();
    openHistoryMenu();

    expect(await screen.findByText('No chats yet')).toBeTruthy();
  });

  it('surfaces a fetch failure instead of silently showing "No chats yet"', async () => {
    fetchHistoryMock.mockRejectedValue(new Error('500'));
    renderHeader();
    openHistoryMenu();

    expect(await screen.findByText("Couldn't load your chat history.")).toBeTruthy();
  });

  it('filters the list by the search box', async () => {
    fetchHistoryMock.mockResolvedValue({
      conversations: [
        historyEntry({ conversationId: 'a', title: 'Add a Remarks column' }),
        historyEntry({ conversationId: 'b', title: 'Freeze the top row' }),
      ],
      nextCursor: null,
    });
    renderHeader();
    openHistoryMenu();

    await screen.findByText('Add a Remarks column');
    fireEvent.change(screen.getByLabelText('Search chats'), { target: { value: 'freeze' } });

    expect(screen.queryByText('Add a Remarks column')).toBeNull();
    expect(screen.queryByText('Freeze the top row')).toBeTruthy();
  });

  it('offers a "load older" control only when the server returned a cursor', async () => {
    fetchHistoryMock.mockResolvedValue({
      conversations: [historyEntry()],
      nextCursor: '2026-08-30T10:00:00.000Z',
    });
    renderHeader();
    openHistoryMenu();

    const more = await screen.findByText('Load older chats');
    fireEvent.click(more);

    await waitFor(() => {
      expect(fetchHistoryMock).toHaveBeenLastCalledWith({
        cursor: '2026-08-30T10:00:00.000Z',
      });
    });
  });

  it('hides the "load older" control on the last page', async () => {
    fetchHistoryMock.mockResolvedValue({
      conversations: [historyEntry()],
      nextCursor: null,
    });
    renderHeader();
    openHistoryMenu();

    await screen.findByText('Add a Remarks column');
    expect(screen.queryByText('Load older chats')).toBeNull();
  });

  it('marks a history row whose thread is already an open tab', async () => {
    fetchHistoryMock.mockResolvedValue({
      conversations: [historyEntry()],
      nextCursor: null,
    });
    const openSession: ChatSession = {
      id: 'sess_1',
      title: 'Add a Remarks column',
      conversationId: 'conv_past',
      turns: [],
      history: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    renderHeader({ sessions: [openSession], activeSessionId: 'sess_other' });
    openHistoryMenu();

    expect(await screen.findByText('Open')).toBeTruthy();
  });
});
