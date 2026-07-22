import { FormatSpec as RichFormatSpec } from '@/action.types';
import { FormatSpec } from '@/types/sheet-actions';

/* global Excel */

function toLegacyFormat(fmt: RichFormatSpec): FormatSpec {
  const borders =
    fmt.borders?.edges.includes('all') || fmt.borders?.edges.includes('outer')
      ? 'outer'
      : fmt.borders?.edges.includes('bottom')
        ? 'bottom'
        : fmt.borders?.style === 'none'
          ? 'none'
          : undefined;

  return {
    bold: fmt.bold,
    italic: fmt.italic,
    fontSize: fmt.fontSize,
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

export function applyFormat(range: Excel.Range, format: FormatSpec): void {
  if (format.bold !== undefined) range.format.font.bold = format.bold;
  if (format.italic !== undefined) range.format.font.italic = format.italic;
  if (format.fontSize !== undefined) range.format.font.size = format.fontSize;
  if (format.fontColor !== undefined) range.format.font.color = format.fontColor;
  if (format.clearFill) {
    range.format.fill.clear();
  } else if (format.fillColor !== undefined) {
    range.format.fill.pattern = 'Solid';
    range.format.fill.color = format.fillColor;
  }
  if (format.numberFormat !== undefined) range.numberFormat = [[format.numberFormat]];
  if (format.horizontalAlignment !== undefined) {
    const map: Record<string, Excel.HorizontalAlignment> = {
      left: 'Left',
      center: 'Center',
      right: 'Right',
    };
    range.format.horizontalAlignment = map[format.horizontalAlignment] ?? 'General';
  }
  if (format.verticalAlignment !== undefined) {
    const map: Record<string, Excel.VerticalAlignment> = {
      top: 'Top',
      middle: 'Center',
      bottom: 'Bottom',
    };
    range.format.verticalAlignment = map[format.verticalAlignment] ?? 'Bottom';
  }
  if (format.wrapText !== undefined) range.format.wrapText = format.wrapText;

  if (format.borders) {
    const borders = range.format.borders;
    const style = 'Continuous' as Excel.BorderStyle;
    if (format.borders === 'all' || format.borders === 'outer') {
      (['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] as Excel.BorderSide[]).forEach((edge) => {
        borders.getItem(edge).style = style;
      });
    }
    if (format.borders === 'bottom') {
      borders.getItem('EdgeBottom').style = style;
    }
    if (format.borders === 'none') {
      (['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] as Excel.BorderSide[]).forEach((edge) => {
        borders.getItem(edge).style = 'None';
      });
    }
  }
}
