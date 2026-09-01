import { describe, expect, it, vi } from 'vitest';
import { applyRichFormat } from './format.handler';
import type { FormatSpec as RichFormatSpec } from '@/action.types';

/**
 * TASKS.md #139 — the header band never painted.
 *
 * `borders` has two incompatible shapes in this codebase (ARCHITECTURE.md AD-7):
 * the rich/shared `{ style, edges: [...] }` and the wire
 * `'all' | 'outer' | 'bottom' | 'none'`. `legacyConverter` casts a wire action's
 * `format` straight to the rich type without converting, so `toLegacyFormat`
 * read `.edges.includes(...)` off a plain string and threw. Because
 * `RichActionEngine.dispatch` catches per-action, the entire FORMAT_RANGE was
 * dropped — silently. Sibling FORMAT_RANGEs with no `borders` (the title, the
 * number formats) applied fine, which is exactly the half-styled workbook that
 * got reported.
 */
function makeRange() {
  const borderItems = new Map<string, { style?: string }>();
  return {
    format: {
      font: {},
      fill: { clear: vi.fn() },
      borders: {
        getItem: (edge: string) => {
          const existing = borderItems.get(edge);
          if (existing) return existing;
          const created: { style?: string } = {};
          borderItems.set(edge, created);
          return created;
        },
      },
    },
    borderItems,
  } as unknown as Excel.Range & { borderItems: Map<string, { style?: string }> };
}

describe('applyRichFormat — borders arrive in two shapes', () => {
  it('does not throw on the wire string shape, and paints the edges', () => {
    const range = makeRange();
    expect(() =>
      applyRichFormat(range, { bold: true, fillColor: '#2F5597', borders: 'all' } as unknown as RichFormatSpec),
    ).not.toThrow();

    // The critical assertion: the rest of the format survived. Previously the
    // throw happened before fill/font were assigned, so the band never painted.
    expect(range.format.font.bold).toBe(true);
    expect(range.format.fill.color).toBe('#2F5597');
    expect([...range.borderItems.keys()]).toEqual(
      expect.arrayContaining(['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight']),
    );
  });

  it('still handles the rich object shape', () => {
    const range = makeRange();
    applyRichFormat(range, {
      bold: true,
      borders: { style: 'thin', edges: ['all'] },
    });
    expect(range.format.font.bold).toBe(true);
    expect(range.borderItems.get('EdgeTop')?.style).toBe('Continuous');
  });

  it('maps the wire "bottom" and "none" strings', () => {
    const bottom = makeRange();
    applyRichFormat(bottom, { borders: 'bottom' } as unknown as RichFormatSpec);
    expect([...bottom.borderItems.keys()]).toEqual(['EdgeBottom']);

    const none = makeRange();
    applyRichFormat(none, { borders: 'none' } as unknown as RichFormatSpec);
    expect(none.borderItems.get('EdgeTop')?.style).toBe('None');
  });

  it('leaves borders untouched when the format omits them', () => {
    const range = makeRange();
    applyRichFormat(range, { bold: true, fontSize: 14 });
    expect(range.borderItems.size).toBe(0);
  });

  it('applies the full header band spec the presentation pass emits', () => {
    const range = makeRange();
    applyRichFormat(range, {
      bold: true,
      fontColor: '#FFFFFF',
      fillColor: '#2F5597',
      horizontalAlignment: 'center',
      verticalAlignment: 'middle',
      wrapText: true,
      borders: 'all',
    } as unknown as RichFormatSpec);

    expect(range.format.font.bold).toBe(true);
    expect(range.format.font.color).toBe('#FFFFFF');
    expect(range.format.fill.color).toBe('#2F5597');
    expect(range.format.horizontalAlignment).toBe('Center');
    // 'middle' is not an Office.js literal — it must be mapped to 'Center'.
    expect(range.format.verticalAlignment).toBe('Center');
    expect(range.format.wrapText).toBe(true);
  });
});
