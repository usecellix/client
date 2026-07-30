import React, { useEffect, useRef, useState } from 'react';
import {
  Archive,
  ChevronRight,
  ExternalLink,
  LogOut,
  MessageSquare,
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

interface PanelHeaderProps {
  sessions: ChatSession[];
  activeSessionId: string | null;
  isWaitingForResponse: boolean;
  onSelectSession: (sessionId: string) => void;
  onCloseSession: (sessionId: string) => void;
  onNewChat: () => void;
}

export const PanelHeader: React.FC<PanelHeaderProps> = ({
  sessions,
  activeSessionId,
  isWaitingForResponse,
  onSelectSession,
  onCloseSession,
  onNewChat,
}) => {
  const { data: session } = useSession();
  const userEmail = session?.user?.email?.trim() || 'Signed in';

  const [historyOpen, setHistoryOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [historyQuery, setHistoryQuery] = useState('');
  const headerRef = useRef<HTMLDivElement>(null);

  const filteredSessions = sessions.filter((chatSession) =>
    chatSession.title.toLowerCase().includes(historyQuery.trim().toLowerCase()),
  );

  const closeAll = () => {
    setHistoryOpen(false);
    setSettingsOpen(false);
  };

  useEffect(() => {
    if (!historyOpen && !settingsOpen) return;

    const handlePointerDown = (event: MouseEvent) => {
      if (headerRef.current && !headerRef.current.contains(event.target as Node)) {
        closeAll();
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    return () => document.removeEventListener('mousedown', handlePointerDown);
  }, [historyOpen, settingsOpen]);

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
              setHistoryOpen((prev) => !prev);
              setSettingsOpen(false);
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
              <div className="cellix-chat-history-section">Today</div>
              <div className="cellix-chat-history-list">
                {filteredSessions.length === 0 ? (
                  <div className="cellix-chat-history-empty">
                    {sessions.length === 0 ? 'No chats yet' : 'No chats found'}
                  </div>
                ) : (
                  filteredSessions.map((chatSession) => {
                    const active = chatSession.id === activeSessionId;
                    return (
                      <div
                        key={chatSession.id}
                        className={`cellix-chat-history-item ${active ? 'active' : ''}`}
                        role="menuitem"
                        tabIndex={0}
                        onClick={() => {
                          onSelectSession(chatSession.id);
                          closeAll();
                        }}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter' || event.key === ' ') {
                            event.preventDefault();
                            onSelectSession(chatSession.id);
                            closeAll();
                          }
                        }}
                        title={chatSession.title}
                      >
                        <MessageSquare size={13} />
                        <span>{chatSession.title}</span>
                        {active && <Pin size={11} className="cellix-chat-history-pin" />}
                        <button
                          type="button"
                          className="cellix-chat-history-delete"
                          title="Remove chat"
                          aria-label="Remove chat"
                          onClick={(event) => {
                            event.stopPropagation();
                            onCloseSession(chatSession.id);
                          }}
                        >
                          <Trash2 size={12} />
                        </button>
                      </div>
                    );
                  })
                )}
              </div>
              <button type="button" className="cellix-chat-history-archived">
                <ChevronRight size={13} />
                <Archive size={13} />
                Archived
              </button>
            </div>
          )}

          <div className="cellix-topbar-settings-wrap">
            <button
              type="button"
              className={`cellix-topbar-icon-btn ${settingsOpen ? 'active' : ''}`}
              onClick={(event) => {
                event.stopPropagation();
                setSettingsOpen((prev) => !prev);
                setHistoryOpen(false);
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
