import React, { useEffect, useRef, useState } from 'react';
import {
  BarChart3,
  ChevronDown,
  Eye,
  EyeOff,
  History,
  Layers,
  MoreHorizontal,
  Plus,
  ScrollText,
} from 'lucide-react';
import { SheetSelector } from '@/components/SheetSelector/SheetSelector';
import { ChangeHistoryPanel } from '@/components/ChangeHistoryPanel/ChangeHistoryPanel';
import { CostDashboardPanel } from '@/components/CostDashboardPanel/CostDashboardPanel';
import { SheetAction } from '@/types/sheet-actions';

type HeaderPanel = 'sheets' | 'history' | 'usage' | 'audit';

interface PanelHeaderProps {
  previewEnabled: boolean;
  onTogglePreview: () => void;
  onClear: () => void;
  hasTurns: boolean;
  conversationId: string | null;
  onRevertChangeSet?: (changeSetId: string, inverseActions: SheetAction[]) => Promise<void>;
  onSheetSelectionChange: (sheets: string[]) => void;
  onCompareSheets: (sheetA: string, sheetB: string) => void;
}

const PANEL_TITLES: Record<HeaderPanel, string> = {
  sheets: 'Sheets',
  history: 'Change history',
  usage: 'Usage',
  audit: 'Audit',
};

export const PanelHeader: React.FC<PanelHeaderProps> = ({
  previewEnabled,
  onTogglePreview,
  onClear,
  hasTurns,
  conversationId,
  onRevertChangeSet,
  onSheetSelectionChange,
  onCompareSheets,
}) => {
  const [openPanel, setOpenPanel] = useState<HeaderPanel | null>(null);
  const [overflowOpen, setOverflowOpen] = useState(false);
  const headerRef = useRef<HTMLDivElement>(null);

  const togglePanel = (panel: HeaderPanel) => {
    setOpenPanel((prev) => (prev === panel ? null : panel));
    setOverflowOpen(false);
  };

  const closeAll = () => {
    setOpenPanel(null);
    setOverflowOpen(false);
  };

  useEffect(() => {
    if (!openPanel && !overflowOpen) return;
    const handle = (event: MouseEvent) => {
      if (headerRef.current && !headerRef.current.contains(event.target as Node)) {
        closeAll();
      }
    };
    document.addEventListener('mousedown', handle);
    return () => document.removeEventListener('mousedown', handle);
  }, [openPanel, overflowOpen]);

  return (
    <div className="cellix-topbar" ref={headerRef}>
      <div className="cellix-topbar-row">
        <div className="cellix-topbar-brand">
          <div className="cellix-topbar-logo">C</div>
          <span className="cellix-topbar-title">Cellix</span>
        </div>

        <div className="cellix-topbar-actions">
          <button
            type="button"
            className={`cellix-header-btn ${openPanel === 'sheets' ? 'active' : ''}`}
            onClick={() => togglePanel('sheets')}
            title="Sheets"
          >
            <Layers size={13} />
            <span className="cellix-header-btn-label">Sheet</span>
            <ChevronDown size={12} />
          </button>

          <button
            type="button"
            onClick={onTogglePreview}
            className={`cellix-header-btn ${previewEnabled ? 'on' : ''}`}
            title="Toggle preview before applying changes"
          >
            {previewEnabled ? <Eye size={13} /> : <EyeOff size={13} />}
            <span className="cellix-header-btn-label">Preview</span>
          </button>

          <div className="cellix-header-overflow-wrap">
            <button
              type="button"
              className={`cellix-header-btn icon-only ${overflowOpen ? 'active' : ''}`}
              onClick={() => {
                setOverflowOpen((prev) => !prev);
                setOpenPanel(null);
              }}
              title="More"
              aria-haspopup="menu"
              aria-expanded={overflowOpen}
            >
              <MoreHorizontal size={15} />
            </button>

            {overflowOpen && (
              <div className="cellix-header-menu" role="menu">
                <button type="button" role="menuitem" onClick={() => togglePanel('history')}>
                  <History size={13} /> Change history
                </button>
                <button type="button" role="menuitem" onClick={() => togglePanel('usage')}>
                  <BarChart3 size={13} /> Usage
                </button>
                <button type="button" role="menuitem" onClick={() => togglePanel('audit')}>
                  <ScrollText size={13} /> Audit
                </button>
              </div>
            )}
          </div>

          {hasTurns && (
            <button
              type="button"
              className="cellix-header-btn icon-only"
              onClick={() => {
                closeAll();
                onClear();
              }}
              title="New chat"
            >
              <Plus size={15} />
            </button>
          )}
        </div>
      </div>

      {openPanel && (
        <div className="cellix-header-popover">
          <div className="cellix-header-popover-head">
            <span>{PANEL_TITLES[openPanel]}</span>
            <button type="button" className="cellix-header-popover-close" onClick={closeAll}>
              ✕
            </button>
          </div>
          <div className="cellix-header-popover-body">
            {openPanel === 'sheets' && (
              <SheetSelector
                embedded
                onSelectionChange={onSheetSelectionChange}
                onCompare={onCompareSheets}
              />
            )}
            {openPanel === 'history' && onRevertChangeSet && (
              <ChangeHistoryPanel
                embedded
                conversationId={conversationId}
                onRevert={onRevertChangeSet}
              />
            )}
            {openPanel === 'usage' && <CostDashboardPanel embedded section="usage" />}
            {openPanel === 'audit' && <CostDashboardPanel embedded section="audit" />}
          </div>
        </div>
      )}
    </div>
  );
};

export default PanelHeader;
