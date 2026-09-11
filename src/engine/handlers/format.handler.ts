import { FormatSpec as RichFormatSpec } from '@/action.types';
import { FormatSpec } from '@/types/sheet-actions';

/* global Excel */

/**
 * `borders` exists in two incompatible shapes — ARCHITECTURE.md AD-7 drift made
 * concrete:
 *
 *   - rich/shared (`Root/shared/action.types.ts`): `{ style, edges: [...] }`
 *   - wire (`client/src/types/sheet-actions.ts`):  `'all' | 'outer' | 'bottom' | 'none'`
 *
 * `legacyConverter` casts a wire action's `format` straight through to the rich
 * type without converting, so a wire `borders: 'all'` arrives here as a plain
 * string. Reading `.edges.includes(...)` off it threw
 * `Cannot read properties of undefined (reading 'includes')`, which
 * `RichActionEngine.dispatch` caught per-action — so the whole FORMAT_RANGE was
 * dropped and the header band silently never painted, while sibling
 * FORMAT_RANGEs carrying no `borders` applied fine. Accept both shapes.
 * TASKS.md #139.
 */
function normalizeBorders(borders: RichFormatSpec['borders'] | FormatSpec['borders']): FormatSpec['borders'] {
  if (!borders) return undefined;
  if (typeof borders === 'string') {
    return borders === 'all' ? 'outer' : borders;
  }
  if (borders.edges?.includes('all') || borders.edges?.includes('outer')) return 'outer';
  if (borders.edges?.includes('bottom')) return 'bottom';
  if (borders.style === 'none') return 'none';
  return undefined;
}

function toLegacyFormat(fmt: RichFormatSpec): FormatSpec {
  const borders = normalizeBorders(fmt.borders);

  return {
    bold: fmt.bold,
    italic: fmt.italic,
    fontSize: fmt.fontSize,
    fontName: fmt.fontName,
    fontColor: fmt.fontColor,
    fillColor: fmt.fillColor,
    numberFormat: fmt.numberFormat,
    horizontalAlignment: fmt.horizontalAlignment,
    verticalAlignment: fmt.verticalAlignment,
    wrapText: fmt.wrapText,
    borders,
  };
}

export function applyRichFormat(range: Excel.Range, format: RichFormatSpec): void {
  applyFormat(range, toLegacyFormat(format));
}

/**
 * `Excel.ConditionalRangeFormat` (used by CONDITIONAL_FORMAT) is a distinct,
 * narrower API from `Excel.RangeFormat` — no `numberFormat` as a 2D array (it's
 * a plain string), no alignment/wrapText, `underline` is a string enum rather
 * than boolean. Deliberately not unified with `applyFormat` above.
 */
export function applyConditionalRangeFormat(
  target: Excel.ConditionalRangeFormat,
  format: RichFormatSpec,
): void {
  if (format.bold !== undefined) target.font.bold = format.bold;
  if (format.italic !== undefined) target.font.italic = format.italic;
  if (format.underline !== undefined) target.font.underline = format.underline ? 'Single' : 'None';
  if (format.fontColor !== undefined) target.font.color = format.fontColor;
  if (format.fillColor !== undefined) target.fill.color = format.fillColor;
  if (format.numberFormat !== undefined) target.numberFormat = format.numberFormat;
}

export function applyFormat(range: Excel.Range, format: FormatSpec): void {
  if (format.bold !== undefined) range.format.font.bold = format.bold;
  if (format.italic !== undefined) range.format.font.italic = format.italic;
  if (format.fontSize !== undefined) range.format.font.size = format.fontSize;
  if (format.fontName !== undefined) range.format.font.name = format.fontName;
  if (format.fontColor !== undefined) range.format.font.color = format.fontColor;
  if (format.clearFill) {
    range.format.fill.clear();
  } else if (format.fillColor !== undefined) {
    range.format.fill.pattern = 'Solid';
    range.format.fill.color = format.fillColor;
  }
  if (format.numberFormat !== undefined) range.numberFormat = [[format.numberFormat]];
  // Office.js types these as the enum widened with its string-literal aliases —
  // index off the property so the literals below stay assignable.
  if (format.horizontalAlignment !== undefined) {
    const map: Record<string, Excel.RangeFormat['horizontalAlignment']> = {
      left: 'Left',
      center: 'Center',
      right: 'Right',
    };
    range.format.horizontalAlignment = map[format.horizontalAlignment] ?? 'General';
  }
  if (format.verticalAlignment !== undefined) {
    const map: Record<string, Excel.RangeFormat['verticalAlignment']> = {
      top: 'Top',
      middle: 'Center',
      bottom: 'Bottom',
    };
    range.format.verticalAlignment = map[format.verticalAlignment] ?? 'Bottom';
  }
  if (format.wrapText !== undefined) range.format.wrapText = format.wrapText;

  if (format.borders) {
    const borders = range.format.borders;
    // Office.js names these BorderLineStyle / BorderIndex — there is no BorderStyle
    // or BorderSide. `as const` keeps the edge literals matching getItem's overload.
    const style: Excel.RangeBorder['style'] = 'Continuous';
    const edges = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] as const;
    if (format.borders === 'all' || format.borders === 'outer') {
      edges.forEach((edge) => {
        borders.getItem(edge).style = style;
      });
    }
    if (format.borders === 'bottom') {
      borders.getItem('EdgeBottom').style = style;
    }
    if (format.borders === 'none') {
      edges.forEach((edge) => {
        borders.getItem(edge).style = 'None';
      });
    }
  }
}
