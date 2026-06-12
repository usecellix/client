import React, { useState, useEffect, useCallback } from 'react';

/* global Excel */

export interface SheetInfo {
  name: string;
  index: number;
  isActive: boolean;
  rowCount?: number;
  colCount?: number;
}

interface SheetSelectorProps {
  onSelectionChange: (selectedSheets: string[]) => void;
  onCompare: (sheetA: string, sheetB: string) => void;
}

export const SheetSelector: React.FC<SheetSelectorProps> = ({
  onSelectionChange,
  onCompare,
}) => {
  const [sheets, setSheets] = useState<SheetInfo[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [compareMode, setCompareMode] = useState(false);
  const [compareA, setCompareA] = useState<string | null>(null);
  const [compareB, setCompareB] = useState<string | null>(null);
  const [isExpanded, setIsExpanded] = useState(false);

  const loadSheets = useCallback(async () => {
    await Excel.run(async (ctx) => {
      const wsList = ctx.workbook.worksheets;
      wsList.load('items/name,items/position,items/visibility');
      const activeWs = ctx.workbook.worksheets.getActiveWorksheet();
      activeWs.load('name');
      await ctx.sync();

      const sheetData: SheetInfo[] = wsList.items
        .filter((ws) => ws.visibility === Excel.SheetVisibility.visible)
        .map((ws) => ({
          name: ws.name,
          index: ws.position,
          isActive: ws.name === activeWs.name,
        }));

      setSheets(sheetData);

      const activeSheet = sheetData.find((s) => s.isActive);
      if (activeSheet) {
        setSelected((prev) => {
          if (prev.size > 0) return prev;
          return new Set([activeSheet.name]);
        });
      }
    });
  }, []);

  useEffect(() => {
    void loadSheets();
  }, [loadSheets]);

  useEffect(() => {
    if (selected.size === 0) return;
    onSelectionChange(Array.from(selected));
  }, [selected, onSelectionChange]);

  const toggleSheet = (sheetName: string) => {
    if (compareMode) {
      if (!compareA) {
        setCompareA(sheetName);
      } else if (!compareB && sheetName !== compareA) {
        setCompareB(sheetName);
      } else {
        setCompareA(sheetName);
        setCompareB(null);
      }
      return;
    }

    const next = new Set(selected);
    if (next.has(sheetName)) {
      if (next.size === 1) return;
      next.delete(sheetName);
    } else {
      next.add(sheetName);
    }
    setSelected(next);
  };

  const handleStartCompare = () => {
    if (compareA && compareB) {
      onCompare(compareA, compareB);
      setCompareMode(false);
      setCompareA(null);
      setCompareB(null);
    }
  };

  const getCompareStatus = (name: string): 'A' | 'B' | null => {
    if (name === compareA) return 'A';
    if (name === compareB) return 'B';
    return null;
  };

  const exitCompareMode = () => {
    setCompareMode(false);
    setCompareA(null);
    setCompareB(null);
  };

  return (
    <div className="cellix-sheet-selector">
      <button
        type="button"
        onClick={() => setIsExpanded(!isExpanded)}
        className="cellix-sheet-selector-header"
      >
        <div className="cellix-sheet-selector-title">
          <span>📋</span>
          <span>Sheets ({selected.size} selected)</span>
        </div>
        <span className="cellix-sheet-selector-chevron">{isExpanded ? '▲' : '▼'}</span>
      </button>

      {isExpanded && (
        <div className="cellix-sheet-selector-body">
          <div className="cellix-sheet-mode-toggle">
            <button
              type="button"
              onClick={exitCompareMode}
              className={`cellix-sheet-mode-btn ${!compareMode ? 'active-select' : ''}`}
            >
              ✓ Select
            </button>
            <button
              type="button"
              onClick={() => setCompareMode(true)}
              className={`cellix-sheet-mode-btn ${compareMode ? 'active-compare' : ''}`}
            >
              ⇄ Compare
            </button>
          </div>

          {compareMode && (
            <div className="cellix-sheet-compare-hint">
              {!compareA
                ? 'Click a sheet to set as Sheet A'
                : !compareB
                  ? `Sheet A: ${compareA} — now pick Sheet B`
                  : `A: ${compareA}  ⇄  B: ${compareB}`}
            </div>
          )}

          {sheets.map((sheet) => {
            const isSelected = selected.has(sheet.name);
            const compareLabel = getCompareStatus(sheet.name);

            return (
              <button
                key={sheet.name}
                type="button"
                onClick={() => toggleSheet(sheet.name)}
                className={`cellix-sheet-row ${
                  isSelected && !compareMode
                    ? 'selected'
                    : compareLabel
                      ? 'compare-picked'
                      : ''
                }`}
              >
                <div className="cellix-sheet-row-left">
                  {compareMode ? (
                    <span
                      className={`cellix-sheet-compare-badge ${
                        compareLabel === 'A'
                          ? 'badge-a'
                          : compareLabel === 'B'
                            ? 'badge-b'
                            : 'badge-empty'
                      }`}
                    >
                      {compareLabel ?? '○'}
                    </span>
                  ) : (
                    <span className={`cellix-sheet-check ${isSelected ? 'checked' : ''}`}>
                      {isSelected && '✓'}
                    </span>
                  )}
                  <span className="cellix-sheet-name">{sheet.name}</span>
                  {sheet.isActive && <span className="cellix-sheet-active-tag">active</span>}
                </div>
              </button>
            );
          })}

          {compareMode && compareA && compareB && (
            <button
              type="button"
              onClick={handleStartCompare}
              className="cellix-sheet-compare-start"
            >
              ⇄ Compare {compareA} vs {compareB}
            </button>
          )}

          <button type="button" onClick={() => void loadSheets()} className="cellix-sheet-refresh">
            ↻ Refresh sheet list
          </button>
        </div>
      )}
    </div>
  );
};

export default SheetSelector;
