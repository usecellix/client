import React, { useEffect, useRef, useState } from 'react';
import {
  Archive,
  BarChart3,
  ChevronRight,
  LogOut,
  MessageSquare,
  MoreHorizontal,
  Pin,
  RotateCcw,
  ScrollText,
  Settings,
  SquarePen,
  Trash2,
  X,
} from 'lucide-react';
import { signOutUser } from '@/auth/useAuth';
import { CostDashboardPanel } from '@/components/CostDashboardPanel/CostDashboardPanel';
import { ChatSession } from '@/types/chatSession';

type HeaderPanel = 'usage' | 'audit';

interface PanelHeaderProps {
  sessions: ChatSession[];
  activeSessionId: string | null;
  isWaitingForResponse: boolean;
  onSelectSession: (sessionId: string) => void;
  onCloseSession: (sessionId: string) => void;
  onNewChat: () => void;
}

const PANEL_TITLES: Record<HeaderPanel, string> = {
  usage: 'Usage',
  audit: 'Audit',
};

export const PanelHeader: React.FC<PanelHeaderProps> = ({
  sessions,
  activeSessionId,
  isWaitingForResponse,
  onSelectSession,
  onCloseSession,
  onNewChat,
}) => {
  const [openPanel, setOpenPanel] = useState<HeaderPanel | null>(null);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [moreOpen, setMoreOpen] = useState(false);
  const [historyQuery, setHistoryQuery] = useState('');
  const headerRef = useRef<HTMLDivElement>(null);

  const filteredSessions = sessions.filter((session) =>
    session.title.toLowerCase().includes(historyQuery.trim().toLowerCase()),
  );

  const closeAll = () => {
    setOpenPanel(null);
    setHistoryOpen(false);
    setMoreOpen(false);
  };

  const togglePanel = (panel: HeaderPanel) => {
    setOpenPanel((prev) => (prev === panel ? null : panel));
    setHistoryOpen(false);
    setMoreOpen(false);
  };

  useEffect(() => {
    if (!openPanel && !historyOpen && !moreOpen) return;

    const handlePointerDown = (event: MouseEvent) => {
      if (headerRef.current && !headerRef.current.contains(event.target as Node)) {
        closeAll();
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    return () => document.removeEventListener('mousedown', handlePointerDown);
  }, [openPanel, historyOpen, moreOpen]);

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
            sessions.map((session) => {
              const active = session.id === activeSessionId;
              const loading = active && isWaitingForResponse;

              return (
                <button
                  key={session.id}
                  type="button"
                  className={`cellix-chat-tab ${active ? 'active' : ''}`}
                  role="tab"
                  aria-selected={active}
                  onClick={() => onSelectSession(session.id)}
                  title={session.title}
                >
                  {loading ? (
                    <span className="cellix-spinner cellix-chat-tab-spinner" />
                  ) : (
                    <MessageSquare size={13} />
                  )}
                  <span className="cellix-chat-tab-title">{session.title}</span>
                  <span
                    role="button"
                    tabIndex={0}
                    className="cellix-chat-tab-close"
                    title="Close chat"
                    onClick={(event) => {
                      event.stopPropagation();
                      onCloseSession(session.id);
                    }}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter' || event.key === ' ') {
                        event.preventDefault();
                        event.stopPropagation();
                        onCloseSession(session.id);
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
          <div className="cellix-topbar-more-wrap">
            <button
              type="button"
              className={`cellix-topbar-icon-btn ${moreOpen ? 'active' : ''}`}
              onClick={() => {
                setMoreOpen((prev) => !prev);
                setOpenPanel(null);
              }}
              title="More"
              aria-label="More"
              aria-haspopup="menu"
              aria-expanded={moreOpen}
            >
              <MoreHorizontal size={16} />
            </button>

            {moreOpen && (
              <div className="cellix-header-menu" role="menu">
                <button type="button" role="menuitem" onClick={() => togglePanel('usage')}>
                  <BarChart3 size={13} /> Usage
                </button>
                <button type="button" role="menuitem" onClick={() => togglePanel('audit')}>
                  <ScrollText size={13} /> Audit
                </button>
              </div>
            )}
          </div>

          <button
            type="button"
            className={`cellix-topbar-icon-btn ${historyOpen ? 'active' : ''}`}
            onClick={() => {
              setHistoryOpen((prev) => !prev);
              setOpenPanel(null);
              setMoreOpen(false);
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
                  filteredSessions.map((session) => {
                    const active = session.id === activeSessionId;
                    return (
                      <div
                        key={session.id}
                        className={`cellix-chat-history-item ${active ? 'active' : ''}`}
                        role="menuitem"
                        tabIndex={0}
                        onClick={() => {
                          onSelectSession(session.id);
                          closeAll();
                        }}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter' || event.key === ' ') {
                            event.preventDefault();
                            onSelectSession(session.id);
                            closeAll();
                          }
                        }}
                        title={session.title}
                      >
                        <MessageSquare size={13} />
                        <span>{session.title}</span>
                        {active && <Pin size={11} className="cellix-chat-history-pin" />}
                        <button
                          type="button"
                          className="cellix-chat-history-delete"
                          title="Remove chat"
                          aria-label="Remove chat"
                          onClick={(event) => {
                            event.stopPropagation();
                            onCloseSession(session.id);
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

          <button
            type="button"
            className={`cellix-topbar-icon-btn ${openPanel === 'usage' ? 'active' : ''}`}
            onClick={() => togglePanel('usage')}
            title="Settings"
            aria-label="Settings"
          >
            <Settings size={16} />
          </button>

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

          <button
            type="button"
            className="cellix-topbar-icon-btn"
            onClick={() => void signOutUser()}
            title="Sign out"
            aria-label="Sign out"
          >
            <LogOut size={16} />
          </button>
        </div>
      </div>

      {openPanel && (
        <div className="cellix-header-popover">
          <div className="cellix-header-popover-head">
            <span>{PANEL_TITLES[openPanel]}</span>
            <button type="button" className="cellix-header-popover-close" onClick={closeAll}>
              x
            </button>
          </div>
          <div className="cellix-header-popover-body">
            {openPanel === 'usage' && <CostDashboardPanel embedded section="usage" />}
            {openPanel === 'audit' && <CostDashboardPanel embedded section="audit" />}
          </div>
        </div>
      )}
    </div>
  );
};

export default PanelHeader;
