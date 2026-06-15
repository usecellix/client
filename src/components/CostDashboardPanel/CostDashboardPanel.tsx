import React, { useCallback, useEffect, useState } from 'react';
import {
  AuditStats,
  downloadAuditExport,
  fetchAuditStats,
} from '@/services/auditService';

function formatUsd(value: number): string {
  return `$${value.toFixed(4)}`;
}

function formatPercent(value: number): string {
  return `${Math.round(value * 100)}%`;
}

interface CostDashboardPanelProps {
  /** When true, render the body directly (no internal toggle) for header popovers. */
  embedded?: boolean;
  /** Which part to render when embedded: usage stats, audit export, or both. */
  section?: 'usage' | 'audit' | 'all';
}

export const CostDashboardPanel: React.FC<CostDashboardPanelProps> = ({
  embedded = false,
  section = 'all',
}) => {
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [stats, setStats] = useState<AuditStats | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [exporting, setExporting] = useState<'json' | 'csv' | null>(null);

  const showBody = embedded || open;
  const showUsage = section === 'usage' || section === 'all';
  const showAudit = section === 'audit' || section === 'all';

  const loadStats = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const result = await fetchAuditStats();
      setStats(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load stats');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (showBody && showUsage && !stats && !loading) {
      void loadStats();
    }
  }, [showBody, showUsage, stats, loading, loadStats]);

  const handleExport = async (format: 'json' | 'csv') => {
    setExporting(format);
    try {
      await downloadAuditExport(format);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Export failed');
    } finally {
      setExporting(null);
    }
  };

  return (
    <div className="cellix-cost-dashboard">
      {!embedded && (
        <button
          type="button"
          className="cellix-cost-dashboard-toggle"
          onClick={() => setOpen((prev) => !prev)}
        >
          {open ? '▾' : '▸'} Usage & audit
        </button>
      )}

      {showBody && (
        <div className="cellix-cost-dashboard-body">
          {error && <p className="cellix-error-text">{error}</p>}

          {showUsage && (
            <>
              {loading && <p className="cellix-preview-summary-text">Loading stats…</p>}
              {stats && (
                <>
                  <div className="cellix-cost-grid">
                    <div className="cellix-cost-stat">
                      <span className="cellix-cost-label">7-day cost</span>
                      <strong>{formatUsd(stats.totalCost)}</strong>
                    </div>
                    <div className="cellix-cost-stat">
                      <span className="cellix-cost-label">LLM calls</span>
                      <strong>{stats.totalCalls}</strong>
                    </div>
                    <div className="cellix-cost-stat">
                      <span className="cellix-cost-label">Tokens</span>
                      <strong>{stats.totalTokens.toLocaleString()}</strong>
                    </div>
                    <div className="cellix-cost-stat">
                      <span className="cellix-cost-label">Success rate</span>
                      <strong>{formatPercent(stats.successRate)}</strong>
                    </div>
                  </div>

                  <div className="cellix-cost-tier-list">
                    {Object.entries(stats.byTier).map(([tier, row]) => (
                      <div key={tier} className="cellix-cost-tier-row">
                        <span>{tier}</span>
                        <span>{row.calls} calls</span>
                        <span>{formatUsd(row.cost)}</span>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </>
          )}

          {showAudit && (
            <div className="cellix-cost-actions">
              <button
                type="button"
                className="cellix-btn-secondary"
                disabled={exporting !== null}
                onClick={() => void handleExport('csv')}
              >
                {exporting === 'csv' ? 'Exporting…' : 'Export audit CSV'}
              </button>
              <button
                type="button"
                className="cellix-btn-secondary"
                disabled={exporting !== null}
                onClick={() => void handleExport('json')}
              >
                {exporting === 'json' ? 'Exporting…' : 'Export audit JSON'}
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default CostDashboardPanel;
