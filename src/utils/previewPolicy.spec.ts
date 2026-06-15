import { describe, expect, it } from 'vitest';
import { requiresExplicitAccept, shouldPreviewActions } from './previewPolicy';

describe('previewPolicy', () => {
  it('requires confirmation for delete sheet', () => {
    expect(requiresExplicitAccept([{ type: 'DELETE_SHEET', sheetName: 'Cellix' }])).toBe(true);
  });

  it('forces preview even when auto-apply is enabled for delete sheet', () => {
    expect(
      shouldPreviewActions([{ type: 'DELETE_SHEET', sheetName: 'Cellix' }], true),
    ).toBe(true);
  });

  it('allows auto-apply for simple cell edits when preview is off', () => {
    expect(
      shouldPreviewActions(
        [{ type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 0, value: 1 }],
        true,
      ),
    ).toBe(false);
  });
});
