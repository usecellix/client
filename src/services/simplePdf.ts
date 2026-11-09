/**
 * Tiny dependency-free PDF writer (text tables only) for the audit-trail PDF export.
 * Built in the browser — nothing goes to a server. Standard Helvetica (WinAnsi), so
 * text is reduced to printable ASCII ("₹" → "Rs."), which the audit trail never needs beyond.
 */

export interface PdfTable {
  title: string;
  subtitle?: string;
  columns: Array<{ header: string; width: number }>;
  rows: string[][];
  footer?: string;
}

const PAGE_W = 842; // A4 landscape, points
const PAGE_H = 595;
const MARGIN = 36;
const FONT = 7.5;
const LINE = 11;

export function pdfSafe(text: string): string {
  return String(text ?? '')
    .replace(/₹/g, 'Rs.')
    .replace(/[–—]/g, '-')
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/…/g, '...')
    .replace(/[·•]/g, '-')
    .replace(/[^\x20-\x7E]/g, '?')
    .replace(/\\/g, '\\\\')
    .replace(/\(/g, '\\(')
    .replace(/\)/g, '\\)');
}

/** Characters that fit a column width at the table font size (Helvetica averages ~0.5 em). */
const fit = (text: string, width: number): string => {
  const max = Math.max(3, Math.floor(width / (FONT * 0.5)));
  return text.length > max ? `${text.slice(0, max - 3)}...` : text;
};

function pageStream(lines: Array<{ x: number; y: number; size: number; text: string; bold?: boolean }>): string {
  return lines
    .map((l) => `BT /${l.bold ? 'F2' : 'F1'} ${l.size} Tf ${l.x.toFixed(1)} ${l.y.toFixed(1)} Td (${pdfSafe(l.text)}) Tj ET`)
    .join('\n');
}

/** Returns the PDF file as a string of 8-bit characters (ASCII only). */
export function buildTablePdf(table: PdfTable, generatedAt = new Date()): string {
  const rowsPerPage = Math.floor((PAGE_H - MARGIN * 2 - 60) / LINE);
  const pages: string[] = [];
  const chunks: string[][][] = [];
  for (let i = 0; i < Math.max(table.rows.length, 1); i += rowsPerPage) chunks.push(table.rows.slice(i, i + rowsPerPage));

  chunks.forEach((chunk, pageIndex) => {
    const lines: Array<{ x: number; y: number; size: number; text: string; bold?: boolean }> = [];
    let y = PAGE_H - MARGIN;
    lines.push({ x: MARGIN, y, size: 13, text: table.title, bold: true });
    y -= 16;
    if (table.subtitle) {
      lines.push({ x: MARGIN, y, size: 8, text: table.subtitle });
      y -= 18;
    }
    let x = MARGIN;
    for (const col of table.columns) {
      lines.push({ x, y, size: FONT, text: fit(col.header, col.width), bold: true });
      x += col.width;
    }
    y -= LINE + 2;
    for (const row of chunk) {
      x = MARGIN;
      table.columns.forEach((col, i) => {
        lines.push({ x, y, size: FONT, text: fit(row[i] ?? '', col.width) });
        x += col.width;
      });
      y -= LINE;
    }
    const foot = `${table.footer ? `${table.footer}   ` : ''}Generated ${generatedAt.toISOString().replace('T', ' ').slice(0, 16)} UTC - page ${pageIndex + 1} of ${chunks.length}`;
    lines.push({ x: MARGIN, y: MARGIN / 2, size: 7, text: foot });
    pages.push(pageStream(lines));
  });

  // Objects: 1 catalog, 2 pages, 3 font, 4 bold font, then per page: page + content.
  const objects: string[] = [];
  const pageIds = pages.map((_, i) => 5 + i * 2);
  objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  objects[2] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pages.length} >>`;
  objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>';
  objects[4] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>';
  pages.forEach((content, i) => {
    const pageId = pageIds[i];
    objects[pageId] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PAGE_W} ${PAGE_H}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ${pageId + 1} 0 R >>`;
    objects[pageId + 1] = `<< /Length ${content.length} >>\nstream\n${content}\nendstream`;
  });

  let out = '%PDF-1.4\n';
  const offsets: number[] = [];
  for (let id = 1; id < objects.length; id += 1) {
    offsets[id] = out.length;
    out += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xref = out.length;
  out += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let id = 1; id < objects.length; id += 1) out += `${String(offsets[id]).padStart(10, '0')} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return out;
}

export function downloadPdf(pdf: string, fileName: string): void {
  const bytes = new Uint8Array(pdf.length);
  for (let i = 0; i < pdf.length; i += 1) bytes[i] = pdf.charCodeAt(i) & 0xff;
  const url = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}
