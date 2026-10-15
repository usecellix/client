// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import { GstReconMissedRowsCard } from '@/components/ConversationPanel/GstReconMissedRowsCard';
import { GstReconMissedRowsBlock } from '@/types/conversationTurn';

const block = (rows: GstReconMissedRowsBlock['rows'], title?: string): GstReconMissedRowsBlock => ({
  id: 'b1',
  type: 'gst_recon_missed_rows',
  rows,
  title,
});

describe('GstReconMissedRowsCard', () => {
  afterEach(() => cleanup());

  it('bank rows: compact pill (amount · row), details in the tooltip, jumps with fullRow=true', () => {
    const onJump = vi.fn();
    render(
      <GstReconMissedRowsCard
        block={block(
          [
            {
              vendorName: 'Books · 2026-05-12 · Orchid Printers — STALE: issued 12-May-2026',
              gstin: '₹15,750 out',
              reason: 'Stale cheques (older than 3 months)',
              explanation: '',
              sheetName: 'Prior BRS (Aug-26)',
              row: 16,
              pillLabel: '₹15,750 · row 16',
              tooltip: 'Books ledger · 2026-05-12 · Orchid Printers',
              selectFullRow: true,
            },
          ],
          'Unreconciled items',
        )}
        onJumpToRow={onJump}
      />,
    );

    const pill = screen.getByRole('button', { name: '₹15,750 · row 16' });
    expect(pill.textContent).toBe('₹15,750 · row 16'); // no narration / date in the capsule
    expect(pill.getAttribute('title')).toContain('Orchid Printers');
    expect(screen.getByText(/Unreconciled items/)).toBeTruthy();

    fireEvent.click(pill);
    expect(onJump).toHaveBeenCalledWith('Prior BRS (Aug-26)', 16, true);
  });

  it('GST rows keep their existing label and single-cell jump', () => {
    const onJump = vi.fn();
    render(
      <GstReconMissedRowsCard
        block={block([
          {
            vendorName: 'Acme Traders',
            gstin: '32AAAAA0000A1Z5',
            reason: 'genuinely_missing',
            explanation: '',
            sheetName: 'Purchase Register',
            row: 7,
          },
        ])}
        onJumpToRow={onJump}
      />,
    );
    const pill = screen.getByRole('button', { name: 'Acme Traders (32AAAAA0000A1Z5)' });
    fireEvent.click(pill);
    expect(onJump).toHaveBeenCalledWith('Purchase Register', 7, undefined);
  });
  it('rows that live only in an attached file are plain text, not a button', () => {
    const onJump = vi.fn();
    render(
      <GstReconMissedRowsCard
        block={block([
          {
            vendorName: 'Bank · 2026-09-04 · CHQ BOOK ISSUE CHG',
            gstin: '₹250 out',
            reason: 'Bank charges not booked',
            explanation: '',
            sheetName: null,
            row: null,
            pillLabel: 'row 18 in attached CSV',
            tooltip: 'In the attached CSV, not in your workbook',
            inFile: true,
          },
          {
            vendorName: 'Books',
            gstin: '',
            reason: 'Stale cheques',
            explanation: '',
            sheetName: 'Bank Ledger',
            row: 8,
            pillLabel: 'Books · row 8',
            selectFullRow: true,
          },
        ])}
        onJumpToRow={onJump}
      />,
    );

    const label = screen.getByText('row 18 in attached CSV');
    expect(label.tagName).toBe('SPAN'); // not a <button>
    expect(label.getAttribute('title')).toMatch(/not in your workbook/);
    expect(screen.queryByRole('button', { name: 'row 18 in attached CSV' })).toBeNull();
    fireEvent.click(label);
    expect(onJump).not.toHaveBeenCalled();

    // the ledger row next to it is still a real button
    fireEvent.click(screen.getByRole('button', { name: 'Books · row 8' }));
    expect(onJump).toHaveBeenCalledWith('Bank Ledger', 8, true);
  });
});
