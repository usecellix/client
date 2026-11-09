import { analyzeFormula, ERROR_VALUES, relativePattern, type LiteralHit } from './formulaAnalysis';
import { columnIndexToLetter } from '@/engine/addressUtils';

/* global Excel */

/**
 * P2.4 / B7 — workbook scanner. Reads every sheet (formulas, values, visibility, hidden
 * rows/columns) and every workbook name, read-only, in row chunks so a 50k-row sheet
 * does not block the task pane. Returns the cross-sheet dependency GRAPH as well
 * (consumed by P3.22 workbook explanation and the colleague's A7.6 dependency map).
 */

export const SCAN_CHUNK_ROWS = 5_000;
/** Hidden row/column detection reads one proxy per row — capped, and the cap is reported. */
export const HIDDEN_ROW_SCAN_CAP = 2_000;
export const HIDDEN_COL_SCAN_CAP = 200;

export interface CellFinding {
  sheet: string;
  address: string;
  formula?: string;
  value?: unknown;
  literals?: LiteralHit[];
}

export interface SheetScan {
  name: string;
  visibility: 'Visible' | 'Hidden' | 'VeryHidden' | string;
  rows: number;
  cols: number;
  formulaCount: number;
  hardcoded: CellFinding[];
  refErrors: CellFinding[];
  errorValues: CellFinding[];
  externalLinks: CellFinding[];
  inconsistent: CellFinding[];
  hiddenRows: number;
  hiddenCols: number;
  hiddenScanCapped: boolean;
  headers: string[];
}

export interface NameScan {
  name: string;
  formula: string;
  broken: boolean;
}

export interface WorkbookScan {
  sheets: SheetScan[];
  names: NameScan[];
  /** sheet → sheets it reads from (formulas only). */
  graph: Record<string, string[]>;
  externalWorkbooks: string[];
}

const addr = (rowIndex: number, colIndex: number) => `${columnIndexToLetter(colIndex)}${rowIndex + 1}`;

/** Analyse one block of formulas/values (pure — used by the reader and by tests). */
export function analyzeBlock(
  sheet: SheetScan,
  formulas: unknown[][],
  values: unknown[][],
  rowOffset: number,
  colOffset: number,
  graph: Map<string, Set<string>>,
  externals: Set<string>,
): void {
  const patterns = new Map<number, Map<string, CellFinding[]>>();
  formulas.forEach((row, r) =>
    row.forEach((f, c) => {
      const absR = rowOffset + r;
      const absC = colOffset + c;
      const value = values[r]?.[c];
      if (typeof value === 'string' && ERROR_VALUES.test(value)) {
        sheet.errorValues.push({ sheet: sheet.name, address: addr(absR, absC), value, formula: typeof f === 'string' && f.startsWith('=') ? f : undefined });
      }
      if (typeof f !== 'string' || !f.startsWith('=')) return;
      sheet.formulaCount += 1;
      const found = analyzeFormula(f);
      const cell: CellFinding = { sheet: sheet.name, address: addr(absR, absC), formula: f };
      if (found.literals.length) sheet.hardcoded.push({ ...cell, literals: found.literals });
      if (found.hasRefError) sheet.refErrors.push(cell);
      if (found.externalWorkbooks.length) {
        sheet.externalLinks.push(cell);
        found.externalWorkbooks.forEach((w) => externals.add(w));
      }
      for (const ref of found.sheetRefs) {
        if (ref.toLowerCase() === sheet.name.toLowerCase()) continue;
        if (!graph.has(sheet.name)) graph.set(sheet.name, new Set());
        graph.get(sheet.name)!.add(ref);
      }
      const pattern = relativePattern(f, absR, absC);
      if (!patterns.has(absC)) patterns.set(absC, new Map());
      const byPattern = patterns.get(absC)!;
      byPattern.set(pattern, [...(byPattern.get(pattern) ?? []), cell]);
    }),
  );
  // A column where one formula differs from a clear majority pattern.
  for (const byPattern of patterns.values()) {
    const groups = [...byPattern.values()].sort((a, b) => b.length - a.length);
    if (groups.length < 2 || groups[0].length < 3) continue;
    for (const odd of groups.slice(1)) if (odd.length === 1) sheet.inconsistent.push(odd[0]);
  }
}

export function emptySheetScan(name: string, visibility: string): SheetScan {
  return {
    name,
    visibility,
    rows: 0,
    cols: 0,
    formulaCount: 0,
    hardcoded: [],
    refErrors: [],
    errorValues: [],
    externalLinks: [],
    inconsistent: [],
    hiddenRows: 0,
    hiddenCols: 0,
    hiddenScanCapped: false,
    headers: [],
  };
}

const yieldToUi = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

/** Read the whole workbook. Read-only — never writes or syncs a change. */
export async function scanWorkbook(): Promise<WorkbookScan> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name,items/visibility');
    const names = ctx.workbook.names;
    names.load('items/name,items/formula,items/type,items/value');
    await ctx.sync();

    const graph = new Map<string, Set<string>>();
    const externals = new Set<string>();
    const out: SheetScan[] = [];

    for (const ws of sheets.items) {
      const scan = emptySheetScan(ws.name, String(ws.visibility));
      out.push(scan);
      const used = ws.getUsedRangeOrNullObject(true);
      used.load(['isNullObject', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
      await ctx.sync();
      if (used.isNullObject) continue;
      scan.rows = used.rowCount;
      scan.cols = used.columnCount;

      for (let start = 0; start < used.rowCount; start += SCAN_CHUNK_ROWS) {
        const count = Math.min(SCAN_CHUNK_ROWS, used.rowCount - start);
        const block = ws.getRangeByIndexes(used.rowIndex + start, used.columnIndex, count, used.columnCount);
        block.load(['formulas', 'values']);
        await ctx.sync();
        if (start === 0) scan.headers = ((block.values?.[0] ?? []) as unknown[]).map((v) => String(v ?? ''));
        analyzeBlock(scan, block.formulas as unknown[][], block.values as unknown[][], used.rowIndex + start, used.columnIndex, graph, externals);
        await yieldToUi();
      }

      const rowsToCheck = Math.min(used.rowIndex + used.rowCount, HIDDEN_ROW_SCAN_CAP);
      const colsToCheck = Math.min(used.columnIndex + used.columnCount, HIDDEN_COL_SCAN_CAP);
      scan.hiddenScanCapped = used.rowIndex + used.rowCount > HIDDEN_ROW_SCAN_CAP || used.columnIndex + used.columnCount > HIDDEN_COL_SCAN_CAP;
      const rowProxies = Array.from({ length: rowsToCheck }, (_, r) => {
        const p = ws.getRangeByIndexes(r, 0, 1, 1).getEntireRow();
        p.load('rowHidden');
        return p;
      });
      const colProxies = Array.from({ length: colsToCheck }, (_, c) => {
        const p = ws.getRangeByIndexes(0, c, 1, 1).getEntireColumn();
        p.load('columnHidden');
        return p;
      });
      await ctx.sync();
      scan.hiddenRows = rowProxies.filter((p) => p.rowHidden).length;
      scan.hiddenCols = colProxies.filter((p) => p.columnHidden).length;
    }

    const nameScans: NameScan[] = names.items.map((n) => {
      const formula = String(n.formula ?? '');
      return { name: n.name, formula, broken: /#REF!/i.test(formula) || /#REF!/i.test(String(n.value ?? '')) };
    });

    return {
      sheets: out,
      names: nameScans,
      graph: Object.fromEntries([...graph.entries()].map(([k, v]) => [k, [...v].sort()])),
      externalWorkbooks: [...externals],
    };
  });
}

// ---------------------------------------------------------------------------- reports

const list = (items: CellFinding[], render: (c: CellFinding) => string, max = 12): string[] => {
  const lines = items.slice(0, max).map((c) => `- ${c.sheet}!${c.address} ${render(c)}`);
  if (items.length > max) lines.push(`- …and ${items.length - max} more`);
  return lines;
};

const literalText = (c: CellFinding) => `\`${c.formula}\` — ${(c.literals ?? []).map((l) => l.text).join(', ')}`;

export function scanReport(scan: WorkbookScan): string {
  const all = <K extends keyof SheetScan>(k: K) => scan.sheets.flatMap((s) => s[k] as unknown as CellFinding[]);
  const hard = all('hardcoded');
  const refs = all('refErrors');
  const errs = all('errorValues').filter((e) => !refs.some((r) => r.sheet === e.sheet && r.address === e.address));
  const ext = all('externalLinks');
  const odd = all('inconsistent');
  const hiddenSheets = scan.sheets.filter((s) => s.visibility !== 'Visible');
  const hiddenRows = scan.sheets.filter((s) => s.hiddenRows || s.hiddenCols);
  const badNames = scan.names.filter((n) => n.broken);
  const formulas = scan.sheets.reduce((n, s) => n + s.formulaCount, 0);
  const lines = [
    `**Workbook scan** — ${scan.sheets.length} sheets, ${formulas} formulas. Read-only: nothing was changed.`,
    '',
    `**Hard-coded numbers in formulas** — ${hard.length}`,
    ...list(hard, literalText),
    `**#REF! / broken references** — ${refs.length + badNames.length}`,
    ...list(refs, (c) => `\`${c.formula}\``),
    ...badNames.map((n) => `- Name **${n.name}** → \`${n.formula}\``),
    `**Other error results** — ${errs.length}`,
    ...list(errs, (c) => `shows ${String(c.value)}${c.formula ? ` (\`${c.formula}\`)` : ''}`),
    `**Links to other workbooks** — ${ext.length}${scan.externalWorkbooks.length ? ` (${scan.externalWorkbooks.join(', ')})` : ''}`,
    ...list(ext, (c) => `\`${c.formula}\``),
    `**Formula breaks the column's pattern** — ${odd.length}`,
    ...list(odd, (c) => `\`${c.formula}\``),
    `**Hidden objects** — ${hiddenSheets.length} hidden sheet(s)${hiddenSheets.length ? `: ${hiddenSheets.map((s) => `${s.name}${s.visibility === 'VeryHidden' ? ' (very hidden)' : ''}`).join(', ')}` : ''}; hidden rows/columns: ${hiddenRows.length ? hiddenRows.map((s) => `${s.name} ${s.hiddenRows} rows / ${s.hiddenCols} cols`).join(', ') : 'none'}`,
  ];
  if (scan.sheets.some((s) => s.hiddenScanCapped)) {
    lines.push(`_Hidden rows/columns checked in the first ${HIDDEN_ROW_SCAN_CAP} rows / ${HIDDEN_COL_SCAN_CAP} columns only._`);
  }
  lines.push('', 'Next: *Find hard-coded numbers in my formulas* · *Check for broken references in this workbook*');
  return lines.join('\n');
}

export function hardcodedReport(scan: WorkbookScan): string {
  const hard = scan.sheets.flatMap((s) => s.hardcoded);
  if (!hard.length) return 'No hard-coded numbers found in formulas (0, 1, 100, 12, 365 and function parameters such as ROUND digits are allowed).';
  return [
    `**${hard.length} formula(s) contain hard-coded numbers** (allowed: 0, 1, 100, 12, 365 and parameters like ROUND's digits):`,
    ...list(hard, literalText, 50),
    '',
    'Fix: ask *Move the hard-coded numbers to input cells* — each constant goes to an **Inputs** sheet and the formula points at it (preview first, Accept to apply, undoable).',
  ].join('\n');
}

export function brokenRefsReport(scan: WorkbookScan): string {
  const refs = scan.sheets.flatMap((s) => s.refErrors);
  const errs = scan.sheets.flatMap((s) => s.errorValues).filter((e) => !refs.some((r) => r.sheet === e.sheet && r.address === e.address));
  const ext = scan.sheets.flatMap((s) => s.externalLinks);
  const badNames = scan.names.filter((n) => n.broken);
  const total = refs.length + errs.length + ext.length + badNames.length;
  if (!total) return 'No broken references, error results, external links or broken names found.';
  return [
    `**Broken references check** — ${total} finding(s). Read-only so far.`,
    `**#REF! in formulas** — ${refs.length}`,
    ...list(refs, (c) => `\`${c.formula}\``),
    `**Names pointing to #REF!** — ${badNames.length}`,
    ...badNames.map((n) => `- **${n.name}** → \`${n.formula}\` (delete it in Formulas → Name Manager)`),
    `**Error results** — ${errs.length}`,
    ...list(errs, (c) => `shows ${String(c.value)}`),
    `**Links to other workbooks** — ${ext.length}`,
    ...list(ext, (c) => `\`${c.formula}\``),
    '',
    'Fix: ask *Break the broken links* — each #REF!/external formula is replaced by the value it shows now (preview first, Accept, undoable).',
  ].join('\n');
}

/** P3.22 — plain-language summary of every sheet (incl. hidden) + cross-sheet links. */
export function explainWorkbook(scan: WorkbookScan): string {
  const purpose = (s: SheetScan): string => {
    const h = s.headers.join(' ').toLowerCase();
    if (!s.rows) return 'empty';
    if (/gstin/.test(h)) return 'GST data (GSTINs)';
    if (/narration|utr|cheque/.test(h)) return 'bank statement / narrations';
    if (/invoice/.test(h) && /date/.test(h)) return 'invoice register';
    if (/asset|depreciation/.test(h)) return 'fixed assets';
    if (/vendor|party|supplier/.test(h)) return 'party / vendor list';
    if (/amount|total/.test(h)) return 'amounts / working';
    if (/sheet|purpose|step/.test(h)) return 'notes / index';
    return 'data table';
  };
  const lines = [`**This workbook has ${scan.sheets.length} sheets** (nothing was changed):`];
  for (const s of scan.sheets) {
    const flags = [
      s.visibility !== 'Visible' ? `**${s.visibility === 'VeryHidden' ? 'very hidden' : 'hidden'}**` : '',
      s.formulaCount ? `${s.formulaCount} formulas` : '',
      s.refErrors.length ? `${s.refErrors.length} #REF!` : '',
    ].filter(Boolean);
    lines.push(`- **${s.name}** — ${purpose(s)}; ${s.rows} rows × ${s.cols} cols${s.headers.filter(Boolean).length ? ` (columns: ${s.headers.filter(Boolean).slice(0, 6).join(', ')}${s.headers.length > 6 ? '…' : ''})` : ''}${flags.length ? ` · ${flags.join(' · ')}` : ''}`);
  }
  const links = Object.entries(scan.graph);
  lines.push('', `**Cross-sheet links** — ${links.length ? links.map(([from, to]) => `${from} → ${to.join(', ')}`).join('; ') : 'none (every sheet stands alone)'}`);
  if (scan.externalWorkbooks.length) lines.push(`**Links to other files:** ${scan.externalWorkbooks.join(', ')}`);
  return lines.join('\n');
}
