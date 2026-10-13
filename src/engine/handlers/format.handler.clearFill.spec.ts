import { describe, expect, it, vi } from 'vitest';
import { applyRichFormat } from './format.handler';
import type { FormatSpec as RichFormatSpec } from '@/action.types';

/**
 * TASKS.md #400 — Revert puts "no fill" back with `clearFill`. The rich
 * FORMAT_RANGE path used to drop that flag on the way to the handler, so a
 * fill could be set by FORMAT_RANGE but never removed by one.
 */
function makeRange() {
  const fill: { clear: ReturnType<typeof vi.fn>; pattern?: string; color?: string } = { clear: vi.fn() };
  return { format: { font: {} as Record<string, unknown>, fill, borders: { getItem: vi.fn(() => ({})) } } };
}

describe('applyRichFormat — clearFill', () => {
  it('removes the fill when clearFill is set', () => {
    const range = makeRange();
    applyRichFormat(range as never, { clearFill: true } as RichFormatSpec);
    expect(range.format.fill.clear).toHaveBeenCalledTimes(1);
    expect(range.format.fill.color).toBeUndefined();
  });

  it('still sets a solid fill for fillColor, without clearing', () => {
    const range = makeRange();
    applyRichFormat(range as never, { fillColor: '#4472C4' } as RichFormatSpec);
    expect(range.format.fill.clear).not.toHaveBeenCalled();
    expect(range.format.fill).toMatchObject({ pattern: 'Solid', color: '#4472C4' });
  });

  it('restores a whole earlier state in one call: bold off, no fill, black text', () => {
    const range = makeRange();
    applyRichFormat(range as never, { bold: false, clearFill: true, fontColor: '#000000' } as RichFormatSpec);
    expect(range.format.font).toMatchObject({ bold: false, color: '#000000' });
    expect(range.format.fill.clear).toHaveBeenCalledTimes(1);
  });
});
