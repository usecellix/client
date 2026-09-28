import React from 'react';
import { GstReconMissedRowsBlock } from '@/types/conversationTurn';

export interface GstReconMissedRowsCardProps {
  block: GstReconMissedRowsBlock;
  onJumpToRow: (sheetName: string, row: number) => void;
}

/** Matches the phrasing already used in the chat summary's reason breakdown (gstReconService.ts). */
const REASON_LABELS: Record<string, string> = {
  blank_counterparty_gstin: 'Blank GSTIN in the register',
  blank_gstin_likely_matched: 'Blank GSTIN, likely matched to a portal invoice',
  blank_taxable_value: 'Taxable value blank',
  ambiguous_rate_slab: 'Ambiguous rate slab — needs CA review',
  gstin_mismatch_same_pan: 'Possible GSTIN mismatch (same vendor, different registration)',
  gstin_not_in_portal: 'Genuinely missing from the portal',
  amount_mismatch: 'Amount differs from the portal record',
  date_mismatch: 'Date differs from the portal record',
  genuinely_missing: 'Genuinely missing from the portal',
};

/** Falls back to a readable Title Case rendering of the raw code for any reason not in the map above. */
function formatReason(reason: string): string {
  if (REASON_LABELS[reason]) return REASON_LABELS[reason];
  return reason
    .split('_')
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

/**
 * Every books row that didn't cleanly match the portal, shown as "<reason>: <link>" —
 * the reason sits as plain inline text, and only the vendor/GSTIN link itself is a
 * capsule pill, jumping straight to that exact cell, always in Purchase Register/Sales
 * Register (the CA's own books), never the GST portal sheet, since this reconciliation
 * is deliberately one-directional. Replaces the old plain-text "Sample missed invoices"
 * bullet list, which both capped at 8 rows and had nothing to click.
 */
export const GstReconMissedRowsCard: React.FC<GstReconMissedRowsCardProps> = ({
  block,
  onJumpToRow,
}) => {
  if (!block.rows.length) return null;

  return (
    <div className="cellix-missed-rows" data-testid="gst-recon-missed-rows-card">
      <div className="cellix-missed-rows-title">Missed rows ({block.rows.length})</div>
      <div className="cellix-missed-rows-pills" data-testid="gst-recon-missed-rows-list">
        {block.rows.map((row, index) => {
          const canJump = Boolean(row.sheetName && row.row);
          const label = `${row.vendorName}${row.gstin ? ` (${row.gstin})` : ' (no GSTIN)'}`;
          return (
            <span
              className="cellix-missed-row-entry"
              key={`${row.sheetName ?? 'unknown'}-${row.row ?? index}`}
            >
              <span className="cellix-missed-row-reason">{formatReason(row.reason)}:</span>
              <button
                type="button"
                className="cellix-missed-row-pill"
                disabled={!canJump}
                onClick={() => canJump && onJumpToRow(row.sheetName as string, row.row as number)}
                title={
                  canJump
                    ? `Jump to ${row.sheetName}, row ${row.row}`
                    : 'No cell reference available for this row'
                }
              >
                {label}
              </button>
            </span>
          );
        })}
      </div>
    </div>
  );
};

export default GstReconMissedRowsCard;
