/**
 * Pure helpers for the client inspection lanes:
 *  - P3.5 move hard-coded numbers to input cells
 *  - P3.6 break broken / external links
 *  - P3.13 + P4.6 lookup #N/A diagnosis and fix
 *  - P3.17 LET rewrite with a round-trip equivalence check
 */
import type { SheetAction } from '@/types/sheet-actions';
import type { CellFinding, WorkbookScan } from './workbookScan';

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Replace the FIRST standalone occurrence of a literal (not part of a ref/number/name). */
export function replaceLiteral(formula: string, literal: string, replacement: string): string {
  const re = new RegExp(`(?<![A-Za-z0-9_$.])${escapeRe(literal)}(?![0-9.%])`);
  return formula.replace(re, replacement);
}

export const INPUTS_SHEET = 'CELLIX Inputs';

export interface PlannedFix {
  actions: SheetAction[];
  lines: string[];
}

/** P3.5 — every hard-coded constant becomes a labelled input cell the formula points at. */
export function planMoveLiteralsToInputs(findings: CellFinding[], existingInputsRows = 0): PlannedFix {
  const inputsOps: Array<{ address: string; value: unknown }> = [];
  const formulaOps = new Map<string, Array<{ address: string; formula: string }>>();
  const lines: string[] = [];
  let row = existingInputsRows > 0 ? existingInputsRows + 1 : 2;
  if (existingInputsRows === 0) {
    inputsOps.push({ address: 'A1', value: 'Input' }, { address: 'B1', value: 'Value' }, { address: 'C1', value: 'Used in' });
  }
  for (const f of findings) {
    let formula = f.formula ?? '';
    for (const lit of f.literals ?? []) {
      const ref = `'${INPUTS_SHEET}'!$B$${row}`;
      const next = replaceLiteral(formula, lit.text.replace(/^[-+]/, ''), ref);
      if (next === formula) continue;
      formula = next;
      inputsOps.push(
        { address: `A${row}`, value: `${f.sheet}!${f.address} constant` },
        { address: `B${row}`, value: lit.value < 0 ? -lit.value : lit.value },
        { address: `C${row}`, value: `${f.sheet}!${f.address}` },
      );
      lines.push(`- ${f.sheet}!${f.address}: \`${f.formula}\` → \`${formula}\` (${lit.text} → ${INPUTS_SHEET}!B${row})`);
      row += 1;
    }
    if (formula !== f.formula) {
      formulaOps.set(f.sheet, [...(formulaOps.get(f.sheet) ?? []), { address: f.address, formula }]);
    }
  }
  const actions: SheetAction[] = [];
  if (inputsOps.length) {
    if (existingInputsRows === 0) actions.push({ type: 'ADD_SHEET', name: INPUTS_SHEET } as SheetAction);
    actions.push({ type: 'BATCH_SET', sheetName: INPUTS_SHEET, operations: inputsOps } as SheetAction);
  }
  for (const [sheetName, ops] of formulaOps) actions.push({ type: 'BATCH_SET', sheetName, operations: ops } as SheetAction);
  return { actions, lines };
}

/** P3.6 — replace #REF! / external-link formulas by the value they show now (marked, not silent). */
export function planBreakLinks(scan: WorkbookScan): PlannedFix {
  const bySheet = new Map<string, Array<{ address: string; value: unknown }>>();
  const lines: string[] = [];
  for (const s of scan.sheets) {
    const cells = [...s.refErrors, ...s.externalLinks.filter((e) => !s.refErrors.some((r) => r.address === e.address))];
    for (const c of cells) {
      const shown = s.errorValues.find((e) => e.address === c.address)?.value;
      const isError = typeof shown === 'string' && shown.startsWith('#');
      const value = isError || shown === undefined ? `#REF! (was ${c.formula})` : shown;
      bySheet.set(s.name, [...(bySheet.get(s.name) ?? []), { address: c.address, value }]);
      lines.push(`- ${s.name}!${c.address}: \`${c.formula}\` → ${isError || shown === undefined ? `text "#REF! (was ${c.formula})" — re-point it by hand` : `value ${String(shown)}`}`);
    }
  }
  return {
    actions: [...bySheet.entries()].map(([sheetName, operations]) => ({ type: 'BATCH_SET', sheetName, operations }) as SheetAction),
    lines,
  };
}

// ---------------------------------------------------------------- lookup diagnosis

export interface LookupDiagnosis {
  column: string;
  naCells: string[];
  lookupRange: string;
  tableFirstColumn: string;
  kind: 'text-vs-number' | 'number-vs-text' | 'spaces' | 'missing' | 'unknown';
  message: string;
  /** Cells to convert to numbers (text-vs-number only). */
  fix?: Array<{ address: string; value: number }>;
}

const colIdx = (letters: string) => {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};
const parseCell = (ref: string) => {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
  return m ? { col: colIdx(m[1]), row: Number(m[2]) - 1 } : null;
};
const parseRange = (ref: string) => {
  const [a, b] = ref.split(':');
  const s = parseCell(a);
  const e = parseCell(b ?? a);
  return s && e ? { r0: s.row, c0: s.col, r1: e.row, c1: e.col } : null;
};
const isNumericText = (v: unknown) => typeof v === 'string' && /^\s*-?\d+(\.\d+)?\s*$/.test(v);

/** Find the lookup args of the first VLOOKUP / XLOOKUP / MATCH in a formula. */
export function lookupArgs(formula: string): { key: string; table: string } | null {
  const m = /\b(VLOOKUP|XLOOKUP|MATCH|HLOOKUP)\(\s*([^,]+?)\s*,\s*([^,)]+)/i.exec(formula);
  if (!m) return null;
  return { key: m[2].replace(/\$/g, ''), table: m[3].replace(/\$/g, '') };
}

export function diagnoseLookupNA(
  formulas: unknown[][],
  values: unknown[][],
  targetColumn?: number,
): LookupDiagnosis | null {
  const width = Math.max(...formulas.map((r) => r.length), 0);
  const cols = targetColumn !== undefined ? [targetColumn] : Array.from({ length: width }, (_, c) => c);
  for (const c of cols) {
    const naRows = values.map((r, i) => (r?.[c] === '#N/A' ? i : -1)).filter((i) => i >= 0);
    if (!naRows.length) continue;
    const formula = String(formulas[naRows[0]]?.[c] ?? '');
    const args = lookupArgs(formula);
    const letter = String.fromCharCode(65 + (c % 26));
    if (!args) {
      return { column: letter, naCells: naRows.map((r) => `${letter}${r + 1}`), lookupRange: '', tableFirstColumn: '', kind: 'unknown', message: `Column ${letter} shows #N/A but is not a lookup formula (\`${formula}\`).` };
    }
    const keyCell = parseCell(args.key);
    const table = parseRange(args.table);
    if (!keyCell || !table) continue;
    const keys = naRows.map((r) => {
      const f = String(formulas[r]?.[c] ?? '');
      const k = parseCell(lookupArgs(f)?.key ?? '');
      return k ? { address: `${String.fromCharCode(65 + (k.col % 26))}${k.row + 1}`, value: values[k.row]?.[k.col] } : null;
    }).filter((k): k is { address: string; value: unknown } => k !== null);
    const firstCol: unknown[] = [];
    for (let r = table.r0; r <= table.r1; r += 1) firstCol.push(values[r]?.[table.c0]);
    const keyLetter = String.fromCharCode(65 + (keyCell.col % 26));
    const tableLetter = String.fromCharCode(65 + (table.c0 % 26));
    const naCells = naRows.map((r) => `${letter}${r + 1}`);
    const base = { column: letter, naCells, lookupRange: `${keyLetter}`, tableFirstColumn: `${tableLetter}` };
    const keysText = keys.filter((k) => isNumericText(k.value));
    const tableNumbers = firstCol.filter((v) => typeof v === 'number');
    if (keysText.length && tableNumbers.length) {
      const matches = keysText.filter((k) => tableNumbers.includes(Number(String(k.value).trim())));
      return {
        ...base,
        kind: 'text-vs-number',
        message: `**Why ${letter} shows #N/A:** the lookup values in column ${keyLetter} are **text** ("${String(keysText[0].value)}") but the codes in column ${tableLetter} are **numbers** (${String(tableNumbers[0])}). Excel never matches text "101" with the number 101. ${matches.length} of ${naCells.length} would match once column ${keyLetter} holds numbers.`,
        fix: keysText.map((k) => ({ address: k.address, value: Number(String(k.value).trim()) })),
      };
    }
    const keysNum = keys.filter((k) => typeof k.value === 'number');
    if (keysNum.length && firstCol.some(isNumericText)) {
      return { ...base, kind: 'number-vs-text', message: `**Why ${letter} shows #N/A:** the lookup values in column ${keyLetter} are **numbers** but the codes in column ${tableLetter} are **text**. Convert column ${tableLetter} to numbers (or wrap the key in TEXT()).` };
    }
    if (keys.some((k) => typeof k.value === 'string' && k.value !== k.value.trim()) || firstCol.some((v) => typeof v === 'string' && v !== v.trim())) {
      return { ...base, kind: 'spaces', message: `**Why ${letter} shows #N/A:** extra spaces in the keys — ask *Clean the spaces in column ${keyLetter}*.` };
    }
    return { ...base, kind: 'missing', message: `**Why ${letter} shows #N/A:** the keys in column ${keyLetter} are simply not in column ${tableLetter} (same type, no extra spaces).` };
  }
  return null;
}

// ---------------------------------------------------------------- LET rewrite

export interface LetRewrite {
  original: string;
  rewritten: string;
  names: Array<{ name: string; expr: string }>;
  equivalent: boolean;
}

const RESERVED = new Set(['TRUE', 'FALSE']);

/** Expand LET names back (inside out) — used to prove the rewrite is equivalent. */
export function expandLet(formula: string): string {
  const m = /^=LET\(([\s\S]*)\)$/.exec(formula);
  if (!m) return formula;
  const parts: string[] = [];
  let depth = 0;
  let buf = '';
  let inStr = false;
  for (const ch of m[1]) {
    if (ch === '"') inStr = !inStr;
    if (!inStr && ch === '(') depth += 1;
    if (!inStr && ch === ')') depth -= 1;
    if (!inStr && depth === 0 && ch === ',') {
      parts.push(buf);
      buf = '';
      continue;
    }
    buf += ch;
  }
  parts.push(buf);
  let body = parts[parts.length - 1];
  for (let i = parts.length - 3; i >= 0; i -= 2) {
    const name = parts[i].trim();
    body = body.replace(new RegExp(`(?<![A-Za-z0-9_."])${escapeRe(name)}(?![A-Za-z0-9_(])`, 'g'), parts[i + 1]);
  }
  return `=${body}`;
}

/**
 * Names every cell reference used ≥ 2 times and every literal-only function call
 * (e.g. DATE(2026,3,31)). Equivalence: expanding the names back must give the
 * original formula character for character — so results are identical on every row.
 */
export function toLet(formula: string, labels: Record<string, string> = {}): LetRewrite | null {
  if (!formula.startsWith('=') || /^=LET\(/i.test(formula)) return null;
  const body = formula.slice(1);
  const names: Array<{ name: string; expr: string }> = [];
  const used = new Set<string>();
  const nameFor = (expr: string, fallback: string) => {
    let base = (labels[expr] ?? fallback).replace(/[^A-Za-z0-9_]/g, '').replace(/^\d/, '_$&') || 'x';
    base = base.charAt(0).toLowerCase() + base.slice(1);
    if (/^[A-Za-z]{1,3}\d+$/.test(base) || RESERVED.has(base.toUpperCase())) base = `_${base}`;
    let n = base;
    let i = 2;
    while (used.has(n)) n = `${base}${i++}`;
    used.add(n);
    return n;
  };
  let rewritten = body;
  const constCalls = [...new Set(body.match(/\b(DATE|TIME)\([\d\s,.-]+\)/gi) ?? [])];
  for (const call of constCalls) {
    const name = nameFor(call, /^DATE/i.test(call) ? 'asOn' : 'k');
    names.push({ name, expr: call });
    rewritten = rewritten.split(call).join(name);
  }
  const refs = analyzeFormulaRefs(rewritten);
  for (const [ref, count] of refs) {
    if (count < 2) continue;
    const name = nameFor(ref, `v${ref.replace(/\$/g, '')}`);
    names.push({ name, expr: ref });
    rewritten = rewritten.replace(new RegExp(`(?<![A-Za-z0-9_$!])${escapeRe(ref)}(?![0-9A-Za-z_(])`, 'g'), name);
  }
  if (!names.length) return null;
  const letFormula = `=LET(${names.map((n) => `${n.name},${n.expr}`).join(',')},${rewritten})`;
  return { original: formula, rewritten: letFormula, names, equivalent: expandLet(letFormula) === formula };
}

function analyzeFormulaRefs(body: string): Map<string, number> {
  const counts = new Map<string, number>();
  const stripped = body.replace(/"[^"]*"/g, '""');
  for (const m of stripped.matchAll(/(?<![A-Za-z0-9_$!])\$?[A-Za-z]{1,3}\$?\d+(?![0-9A-Za-z_(:])/g)) {
    counts.set(m[0], (counts.get(m[0]) ?? 0) + 1);
  }
  return counts;
}
