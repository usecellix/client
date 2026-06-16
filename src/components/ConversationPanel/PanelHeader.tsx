import React, { useEffect, useRef, useState } from 'react';
import {
  Archive,
  BarChart3,
  ChevronRight,
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
import { CostDashboardPanel } from '@/components/CostDashboardPanel/CostDashboardPanel';
import { ConversationTurn } from '@/types/conversationTurn';

type HeaderPanel = 'usage' | 'audit';

interface PanelHeaderProps {
  turns: ConversationTurn[];
  activeTurnId: string | null;
  draftChatOpen: boolean;
  isWaitingForResponse: boolean;
  onSelectTurn: (turnId: string) => void;
  onCloseTurn: (turnId: string) => void;
  onNewChat: () => void;
  onCloseDraftChat: () => void;
}

const PANEL_TITLES: Record<HeaderPanel, string> = {
  usage: 'Usage',
  audit: 'Audit',
};

export const PanelHeader: React.FC<PanelHeaderProps> = ({
  turns,
  activeTurnId,
  draftChatOpen,
  isWaitingForResponse,
  onSelectTurn,
  onCloseTurn,
  onNewChat,
  onCloseDraftChat,
}) => {
  const [openPanel, setOpenPanel] = useState<HeaderPanel | null>(null);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [moreOpen, setMoreOpen] = useState(false);
  const [historyQuery, setHistoryQuery] = useState('');
  const headerRef = useRef<HTMLDivElement>(null);

  const filteredTurns = turns.filter((turn) =>
    turn.userMessage.toLowerCase().includes(historyQuery.trim().toLowerCase()),
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
          {turns.length === 0 ? (
            <button
              type="button"
              className="cellix-chat-tab active cellix-new-chat-tab"
              role="tab"
              aria-selected
            >
              <MessageSquare size={14} />
              <span className="cellix-chat-tab-title">New chat</span>
              <span
                role="button"
                tabIndex={0}
                className="cellix-chat-tab-close"
                title="Close chat"
                onClick={(event) => {
                  event.stopPropagation();
                  onCloseDraftChat();
                }}
                onKeyDown={(event) => {
                  if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault();
                    event.stopPropagation();
                    onCloseDraftChat();
                  }
                }}
              >
                <X size={12} />
              </span>
            </button>
          ) : (
            turns.map((turn) => {
              const active = turn.id === activeTurnId;
              const loading = active && isWaitingForResponse;

              return (
                <button
                  key={turn.id}
                  type="button"
                  className={`cellix-chat-tab ${active ? 'active' : ''}`}
                  role="tab"
                  aria-selected={active}
                  onClick={() => onSelectTurn(turn.id)}
                  title={turn.userMessage}
                >
                  {loading ? (
                    <span className="cellix-spinner cellix-chat-tab-spinner" />
                  ) : (
                    <MessageSquare size={13} />
                  )}
                  <span className="cellix-chat-tab-title">{turn.tabLabel}</span>
                  <span
                    role="button"
                    tabIndex={0}
                    className="cellix-chat-tab-close"
                    title="Close chat"
                    onClick={(event) => {
                      event.stopPropagation();
                      onCloseTurn(turn.id);
                    }}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter' || event.key === ' ') {
                        event.preventDefault();
                        event.stopPropagation();
                        onCloseTurn(turn.id);
                      }
                    }}
                  >
                    <X size={12} />
                  </span>
                </button>
              );
            })
          )}
          {turns.length > 0 && draftChatOpen && (
            <button
              type="button"
              className={`cellix-chat-tab cellix-new-chat-tab ${activeTurnId === null ? 'active' : ''}`}
              onClick={() => {
                closeAll();
                onNewChat();
              }}
              title="New chat"
            >
              <MessageSquare size={13} />
              <span className="cellix-chat-tab-title">New chat</span>
              <span
                role="button"
                tabIndex={0}
                className="cellix-chat-tab-close"
                title="Close chat"
                onClick={(event) => {
                  event.stopPropagation();
                  onCloseDraftChat();
                }}
                onKeyDown={(event) => {
                  if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault();
                    event.stopPropagation();
                    onCloseDraftChat();
                  }
                }}
              >
                <X size={12} />
              </span>
            </button>
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
                {filteredTurns.length === 0 ? (
                  <div className="cellix-chat-history-empty">
                    {turns.length === 0 ? 'No chats yet' : 'No chats found'}
                  </div>
                ) : (
                  filteredTurns.map((turn) => {
                    const active = turn.id === activeTurnId;
                    return (
                      <div
                        key={turn.id}
                        className={`cellix-chat-history-item ${active ? 'active' : ''}`}
                        role="menuitem"
                        tabIndex={0}
                        onClick={() => {
                          onSelectTurn(turn.id);
                          closeAll();
                        }}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter' || event.key === ' ') {
                            event.preventDefault();
                            onSelectTurn(turn.id);
                            closeAll();
                          }
                        }}
                        title={turn.userMessage}
                      >
                        <MessageSquare size={13} />
                        <span>{turn.tabLabel}</span>
                        {active && <Pin size={11} className="cellix-chat-history-pin" />}
                        <button
                          type="button"
                          className="cellix-chat-history-delete"
                          title="Remove chat"
                          aria-label="Remove chat"
                          onClick={(event) => {
                            event.stopPropagation();
                            onCloseTurn(turn.id);
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
