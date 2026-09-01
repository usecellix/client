import { describe, expect, it } from 'vitest';
import {
  blockedActionsAreDataWrites,
  computeSheetLayout,
  HEADER_ROW,
  isHeaderFormatCorrectionMessage,
  sanitizeActions,
} from '@/utils/actionGuard';
import { SheetAction } from '@/types/sheet-actions';

describe('actionGuard header cosmetics (Spec 24 follow-up)', () => {
  const layout = computeSheetLayout([
    ['Date', 'Supplier', 'Amount'],
    ['2024-01-01', 'Acme', 100],
  ]);

  it('allows FORMAT_RANGE on the header row (does not force row-placement clarify)', () => {
    const actions: SheetAction[] = [
      {
        type: 'FORMAT_RANGE',
        sheetName: 'Purchase Register',
        row: HEADER_ROW,
        col: 0,
        rowCount: 1,
        colCount: 3,
        format: { fillColor: '#C6EFCE' },
      },
    ];

    const result = sanitizeActions(actions, layout);
    expect(result.requiresClarification).toBe(false);
    expect(result.actions).toHaveLength(1);
    expect(result.actions[0]?.type).toBe('FORMAT_RANGE');
    expect(result.blocked).toHaveLength(0);
  });

  it('allows HIGHLIGHT_CELL on the header row', () => {
    const result = sanitizeActions(
      [{ type: 'HIGHLIGHT_CELL', row: 0, col: 0, color: '#C6EFCE' }],
      layout,
    );
    expect(result.actions).toHaveLength(1);
    expect(result.requiresClarification).toBe(false);
  });

  it('still blocks header SET_CELL value overwrites and can convert / clarify as data writes', () => {
    const result = sanitizeActions(
      [
        { type: 'SET_CELL', row: 0, col: 0, value: 'New Date label' },
        { type: 'SET_CELL', row: 0, col: 1, value: 'New Supplier' },
      ],
      layout,
    );
    // Converted to ADD_ROW — not left as header writes
    expect(result.actions.some((a) => a.type === 'ADD_ROW')).toBe(true);
    expect(result.actions.every((a) => a.row !== 0 || a.type === 'ADD_ROW')).toBe(true);
  });

  it('blockedActionsAreDataWrites is false for pure format blocks', () => {
    expect(
      blockedActionsAreDataWrites([
        { type: 'FORMAT_RANGE', row: 0, col: 0, rowCount: 1, colCount: 3, format: { bold: true } },
      ]),
    ).toBe(false);
    expect(
      blockedActionsAreDataWrites([{ type: 'SET_CELL', row: 0, col: 0, value: 'x' }]),
    ).toBe(true);
  });

  it('detects free-text correction that rejects new-row clarify', () => {
    expect(
      isHeaderFormatCorrectionMessage(
        'not new row, the existing row, add bg color to header row',
      ),
    ).toBe(true);
    expect(isHeaderFormatCorrectionMessage('After the last row (recommended)')).toBe(false);
  });
});
