import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  Check,
  ExternalLink,
  Flag,
  LogOut,
  MessageSquare,
  Pencil,
  Pin,
  RotateCcw,
  Settings,
  Sparkles,
  SquarePen,
  Trash2,
  User,
  X,
} from 'lucide-react';
import { useSession } from '@/auth/auth-client';
import { signOutUser } from '@/auth/useAuth';
import { ChatSession } from '@/types/chatSession';
import { CheckpointPanel } from '@/components/CheckpointPanel/CheckpointPanel';
import { RestoreResult } from '@/types/checkpoint';
import {
  ConversationSummary,
  fetchConversationHistory,
  renameConversation,
} from '@/services/conversationHistoryService';
import {
  dedupeConversations,
  groupConversationsByRecency,
} from '@/utils/conversationHistoryGrouping';

interface PanelHeaderProps {
  sessions: ChatSession[];
  activeSessionId: string | null;
  isWaitingForResponse: boolean;
  onSelectSession: (sessionId: string) => void;
  onCloseSession: (sessionId: string) => void;
  /** Rename an open tab, and its server conversation if it has one (TASKS.md #177). */
  onRenameSession: (sessionId: string, title: string) => void;
  /** Delete an open tab, and its server conversation if it has one (TASKS.md #177). */
  onDeleteSession: (sessionId: string) => Promise<void>;
  /** Delete a history row that isn't necessarily an open tab (TASKS.md #177). */
  onDeleteHistoryConversation: (conversationId: string) => Promise<void>;
  onNewChat: () => void;
  /**
   * Open a past conversation from server-backed history (TASKS.md #172).
   * Resolves false on failure so the menu can say so instead of quietly
   * swapping in an empty thread.
   */
  onOpenHistoryConversation: (conversationId: string) => Promise<boolean>;
  /** Checkpoints icon only makes sense once a conversation exists. Change
   *  History was removed from here — reverting a specific action is now done
   *  inline on that message (TurnRenderer's own Revert icon) instead of via
   *  a separate browsable panel. */
  showCheckpointsButton?: boolean;
  workbookId?: string;
  conversationId: string | null;
  onRestoreCheckpoint: (result: RestoreResult) => Promise<void>;
}

export const PanelHeader: React.FC<PanelHeaderProps> = ({
  sessions,
  activeSessionId,
  isWaitingForResponse,
  onSelectSession,
  onCloseSession,
  onRenameSession,
  onDeleteSession,
  onDeleteHistoryConversation,
  onNewChat,
  onOpenHistoryConversation,
  showCheckpointsButton = false,
  workbookId,
  conversationId,
  onRestoreCheckpoint,
}) => {
  const { data: session } = useSession();
  const userEmail = session?.user?.email?.trim() || 'Signed in';

  const [historyOpen, setHistoryOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [checkpointsOpen, setCheckpointsOpen] = useState(false);
  const [historyQuery, setHistoryQuery] = useState('');
  const [history, setHistory] = useState<ConversationSummary[]>([]);
  const [historyCursor, setHistoryCursor] = useState<string | null>(null);
  const [historyLoading, setHistoryLoading] = useState(false);
  const [historyError, setHistoryError] = useState<string | null>(null);
  const [openingId, setOpeningId] = useState<string | null>(null);
  /** `{ kind, id }` of the row currently in inline rename edit (TASKS.md #177). */
  const [renaming, setRenaming] = useState<{ kind: 'tab' | 'history'; id: string } | null>(null);
  const [renameDraft, setRenameDraft] = useState('');
  /** `{ kind, id }` of the row awaiting a second click to confirm deletion. */
  const [confirmingDelete, setConfirmingDelete] = useState<{
    kind: 'tab' | 'history';
    id: string;
  } | null>(null);
  const [deletingId, setDeletingId] = useState<string | null>(null);
  const headerRef = useRef<HTMLDivElement>(null);
  const renameInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (renaming) {
      renameInputRef.current?.focus();
      renameInputRef.current?.select();
    }
  }, [renaming]);

  const startRename = (kind: 'tab' | 'history', id: string, currentTitle: string) => {
    setConfirmingDelete(null);
    setRenaming({ kind, id });
    setRenameDraft(currentTitle);
  };

  const cancelRename = () => {
    setRenaming(null);
    setRenameDraft('');
  };

  const commitRename = () => {
    if (!renaming) return;
    const trimmed = renameDraft.trim();
    if (trimmed) {
      if (renaming.kind === 'tab') {
        onRenameSession(renaming.id, trimmed);
      } else {
        // A history row with no open tab has no local session to rename, so
        // the server call goes straight through the history endpoint rather
        // than useConversation's tab-oriented onRenameSession.
        void renameConversation(renaming.id, trimmed)
          .then(() => {
            setHistory((prev) =>
              prev.map((entry) =>
                entry.conversationId === renaming.id ? { ...entry, title: trimmed } : entry,
              ),
            );
          })
          .catch((error) => {
            console.warn('[Cellix] Failed to rename chat:', error);
            setHistoryError("Couldn't rename that chat.");
          });
      }
    }
    cancelRename();
  };

  /** Two-click confirm — a destructive, unrecoverable action gets no single-click trigger. */
  const requestDelete = (kind: 'tab' | 'history', id: string) => {
    if (confirmingDelete?.kind === kind && confirmingDelete.id === id) {
      void commitDelete(kind, id);
      return;
    }
    setRenaming(null);
    setConfirmingDelete({ kind, id });
  };

  const commitDelete = async (kind: 'tab' | 'history', id: string) => {
    setConfirmingDelete(null);
    setDeletingId(id);
    try {
      if (kind === 'tab') {
        await onDeleteSession(id);
        // The tab may have been the source of a history row too — drop it
        // from the loaded list so the menu doesn't show a stale entry
        // pointing at a conversation that's now gone.
        const removedConversationId = sessions.find((s) => s.id === id)?.conversationId;
        if (removedConversationId) {
          setHistory((prev) =>
            prev.filter((entry) => entry.conversationId !== removedConversationId),
          );
        }
      } else {
        await onDeleteHistoryConversation(id);
        setHistory((prev) => prev.filter((entry) => entry.conversationId !== id));
      }
    } catch (error) {
      console.warn('[Cellix] Failed to delete chat:', error);
      setHistoryError("Couldn't delete that chat.");
    } finally {
      setDeletingId(null);
    }
  };

  /**
   * Fetched on open rather than on mount — history is behind a menu almost
   * nobody opens on every session, so paying for the request up front would be
   * a round trip most users never look at.
   */
  const loadHistory = useCallback(async (cursor?: string) => {
    setHistoryLoading(true);
    setHistoryError(null);
    try {
      const page = await fetchConversationHistory(cursor ? { cursor } : undefined);
      setHistory((prev) =>
        cursor ? dedupeConversations([...prev, ...page.conversations]) : page.conversations,
      );
      setHistoryCursor(page.nextCursor);
    } catch (error) {
      console.warn('[Cellix] Failed to load chat history:', error);
      setHistoryError("Couldn't load your chat history.");
    } finally {
      setHistoryLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!historyOpen) return;
    void loadHistory();
  }, [historyOpen, loadHistory]);

  const handleOpenHistoryConversation = async (targetId: string) => {
    setOpeningId(targetId);
    try {
      const opened = await onOpenHistoryConversation(targetId);
      if (opened) {
        closeAll();
      } else {
        setHistoryError("Couldn't open that chat.");
      }
    } finally {
      setOpeningId(null);
    }
  };

  const query = historyQuery.trim().toLowerCase();
  const filteredHistory = query
    ? history.filter(
        (entry) =>
          entry.title.toLowerCase().includes(query) ||
          entry.lastMessage.toLowerCase().includes(query),
      )
    : history;
  const historyGroups = groupConversationsByRecency(filteredHistory);

  /** conversationIds already open as tabs — marked so they read as "current". */
  const openConversationIds = new Set(
    sessions.map((chatSession) => chatSession.conversationId).filter(Boolean) as string[],
  );
  const activeConversationId =
    sessions.find((chatSession) => chatSession.id === activeSessionId)?.conversationId ?? null;

  const closeAll = () => {
    setHistoryOpen(false);
    setSettingsOpen(false);
    setCheckpointsOpen(false);
    setRenaming(null);
    setConfirmingDelete(null);
  };

  useEffect(() => {
    if (!historyOpen && !settingsOpen && !checkpointsOpen) return;

    const handlePointerDown = (event: MouseEvent) => {
      if (headerRef.current && !headerRef.current.contains(event.target as Node)) {
        closeAll();
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    return () => document.removeEventListener('mousedown', handlePointerDown);
  }, [historyOpen, settingsOpen, checkpointsOpen]);

  return (
    <div className="cellix-topbar" ref={headerRef}>
      <div className="cellix-topbar-row">
        <div className="cellix-chat-tab-strip" role="tablist" aria-label="Chats">
          {sessions.length === 0 ? (
            <button
              type="button"
              className="cellix-chat-tab active cellix-new-chat-tab"
              role="tab"
              aria-selected
            >
              <MessageSquare size={14} />
              <span className="cellix-chat-tab-title">New chat</span>
            </button>
          ) : (
            sessions.map((chatSession) => {
              const active = chatSession.id === activeSessionId;
              const loading = active && isWaitingForResponse;

              return (
                <button
                  key={chatSession.id}
                  type="button"
                  className={`cellix-chat-tab ${active ? 'active' : ''}`}
                  role="tab"
                  aria-selected={active}
                  onClick={() => onSelectSession(chatSession.id)}
                  title={chatSession.title}
                >
                  {loading ? (
                    <span className="cellix-spinner cellix-chat-tab-spinner" />
                  ) : (
                    <MessageSquare size={13} />
                  )}
                  <span className="cellix-chat-tab-title">{chatSession.title}</span>
                  <span
                    role="button"
                    tabIndex={0}
                    className="cellix-chat-tab-close"
                    title="Close chat"
                    onClick={(event) => {
                      event.stopPropagation();
                      onCloseSession(chatSession.id);
                    }}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter' || event.key === ' ') {
                        event.preventDefault();
                        event.stopPropagation();
                        onCloseSession(chatSession.id);
                      }
                    }}
                  >
                    <X size={12} />
                  </span>
                </button>
              );
            })
          )}
        </div>

        <div className="cellix-topbar-icons">
          <button
            type="button"
            className={`cellix-topbar-icon-btn ${historyOpen ? 'active' : ''}`}
            onClick={() => {
              const next = !historyOpen;
              closeAll();
              setHistoryOpen(next);
            }}
            title="Chat history"
            aria-label="Chat history"
            aria-haspopup="menu"
            aria-expanded={historyOpen}
          >
            <RotateCcw size={16} />
          </button>

          {historyOpen && (
            <div className="cellix-chat-history-menu" role="menu" aria-label="Chat history">
              <input
                className="cellix-chat-history-search"
                value={historyQuery}
                onChange={(event) => setHistoryQuery(event.target.value)}
                placeholder="Search chats..."
                aria-label="Search chats"
              />

              {historyError && (
                <div className="cellix-chat-history-error" role="alert">
                  {historyError}
                  <button
                    type="button"
                    className="cellix-chat-history-retry"
                    onClick={() => void loadHistory()}
                  >
                    Retry
                  </button>
                </div>
              )}

              <div className="cellix-chat-history-list">
                {historyLoading && history.length === 0 ? (
                  <div className="cellix-chat-history-empty">Loading chats…</div>
                ) : historyGroups.length === 0 ? (
                  <div className="cellix-chat-history-empty">
                    {history.length === 0 ? 'No chats yet' : 'No chats found'}
                  </div>
                ) : (
                  historyGroups.map((group) => (
                    <React.Fragment key={group.label}>
                      <div className="cellix-chat-history-section">{group.label}</div>
                      {group.conversations.map((entry) => {
                        const active = entry.conversationId === activeConversationId;
                        const opening = openingId === entry.conversationId;
                        const open = openConversationIds.has(entry.conversationId);
                        return (
                          <div
                            key={entry.conversationId}
                            className={`cellix-chat-history-item ${active ? 'active' : ''}`}
                            role="menuitem"
                            tabIndex={0}
                            aria-busy={opening}
                            onClick={() => void handleOpenHistoryConversation(entry.conversationId)}
                            onKeyDown={(event) => {
                              if (event.key === 'Enter' || event.key === ' ') {
                                event.preventDefault();
                                void handleOpenHistoryConversation(entry.conversationId);
                              }
                            }}
                            title={entry.title}
                          >
                            {opening ? (
                              <span className="cellix-spinner cellix-chat-tab-spinner" />
                            ) : (
                              <MessageSquare size={13} />
                            )}
                            <span>{entry.title}</span>
                            {active ? (
                              <Pin size={11} className="cellix-chat-history-pin" />
                            ) : open ? (
                              <span className="cellix-chat-history-badge">Open</span>
                            ) : null}
                          </div>
                        );
                      })}
                    </React.Fragment>
                  ))
                )}
              </div>

              {historyCursor && (
                <button
                  type="button"
                  className="cellix-chat-history-more"
                  disabled={historyLoading}
                  onClick={() => void loadHistory(historyCursor)}
                >
                  {historyLoading ? 'Loading…' : 'Load older chats'}
                </button>
              )}
            </div>
          )}

          {showCheckpointsButton && (
            <div className="cellix-topbar-checkpoints-wrap">
              <button
                type="button"
                className={`cellix-topbar-icon-btn ${checkpointsOpen ? 'active' : ''}`}
                onClick={() => {
                  const next = !checkpointsOpen;
                  closeAll();
                  setCheckpointsOpen(next);
                }}
                title="Checkpoints"
                aria-label="Checkpoints"
                aria-haspopup="menu"
                aria-expanded={checkpointsOpen}
              >
                <Flag size={16} />
              </button>

              {checkpointsOpen && (
                <div className="cellix-checkpoint-menu" role="menu" aria-label="Checkpoints">
                  <CheckpointPanel
                    workbookId={workbookId}
                    conversationId={conversationId}
                    onRestore={onRestoreCheckpoint}
                    embedded
                  />
                </div>
              )}
            </div>
          )}

          <div className="cellix-topbar-settings-wrap">
            <button
              type="button"
              className={`cellix-topbar-icon-btn ${settingsOpen ? 'active' : ''}`}
              onClick={(event) => {
                event.stopPropagation();
                const next = !settingsOpen;
                closeAll();
                setSettingsOpen(next);
              }}
              title="Settings"
              aria-label="Settings"
              aria-haspopup="menu"
              aria-expanded={settingsOpen}
            >
              <Settings size={16} />
            </button>
          </div>

          <button
            type="button"
            className="cellix-topbar-icon-btn"
            onClick={() => {
              closeAll();
              onNewChat();
            }}
            title="New chat"
            aria-label="New chat"
          >
            <SquarePen size={16} />
          </button>
        </div>
      </div>

      {settingsOpen && (
        <div className="cellix-settings-menu" role="menu" aria-label="Settings">
          <div className="cellix-settings-menu-section">Account</div>
          <div className="cellix-settings-menu-meta" title={userEmail}>
            <User size={13} />
            <span>{userEmail}</span>
          </div>

          <button
            type="button"
            className="cellix-settings-menu-item"
            role="menuitem"
            onClick={() => {
              closeAll();
            }}
          >
            <Sparkles size={13} />
            <span>Upgrade the plan</span>
            <ExternalLink size={12} className="cellix-settings-menu-trailing" />
          </button>

          <div className="cellix-settings-menu-divider" />

          <button
            type="button"
            className="cellix-settings-menu-item"
            role="menuitem"
            onClick={() => {
              closeAll();
            }}
          >
            <Settings size={13} />
            <span>Cellix settings</span>
          </button>

          <div className="cellix-settings-menu-divider" />

          <button
            type="button"
            className="cellix-settings-menu-item cellix-settings-menu-item--danger"
            role="menuitem"
            onClick={() => {
              closeAll();
              void signOutUser();
            }}
          >
            <LogOut size={13} />
            <span>Log out</span>
          </button>
        </div>
      )}
    </div>
  );
};

export default PanelHeader;
