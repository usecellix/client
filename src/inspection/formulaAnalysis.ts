/**
 * Pure formula analysis for the workbook scanner (P2.4 / B7), the hard-coded value finder
 * (P3.5 / A7.2) and the broken-reference finder (P3.6 / A7.3). No Office.js here — the
 * reader in workbookScan.ts feeds it strings, so everything is unit-testable.
 */

/** Constants that are almost never "assumptions": 0, 1, 100, 12 (months), 365 (days). */
export const LITERAL_ALLOWLIST = new Set([0, 1, 100, 12, 365]);

/** Functions whose numeric argument at this position is a parameter, not an assumption. */
const PARAMETER_ARGS: Record<string, number[]> = {
  ROUND: [1],
  ROUNDUP: [1],
  ROUNDDOWN: [1],
  MROUND: [1],
  VLOOKUP: [2, 3],
  HLOOKUP: [2, 3],
  INDEX: [1, 2],
  MATCH: [2],
  XLOOKUP: [4, 5],
  LEFT: [1],
  RIGHT: [1],
  MID: [1, 2],
  OFFSET: [1, 2, 3, 4],
  SUBTOTAL: [0],
  AGGREGATE: [0, 1],
  WEEKDAY: [1],
  EOMONTH: [1],
  EDATE: [1],
  DATE: [0, 1, 2],
  TEXT: [1],
  CHOOSE: [0],
  LARGE: [1],
  SMALL: [1],
  RANK: [2],
};

export interface LiteralHit {
  text: string;
  value: number;
}

export interface FormulaFinding {
  literals: LiteralHit[];
  hasRefError: boolean;
  externalWorkbooks: string[];
  sheetRefs: string[];
}

interface Token {
  kind: 'num' | 'str' | 'ref' | 'func' | 'op' | 'open' | 'close' | 'comma' | 'other';
  text: string;
}

function tokenize(formula: string): Token[] {
  const src = formula.startsWith('=') ? formula.slice(1) : formula;
  const out: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < src.length && !(src[j] === '"' && src[j + 1] !== '"')) j += src[j] === '"' ? 2 : 1;
      out.push({ kind: 'str', text: src.slice(i, j + 1) });
      i = j + 1;
      continue;
    }
    if (ch === "'") {
      // quoted sheet name: 'Sheet 1'!A1
      let j = i + 1;
      while (j < src.length && !(src[j] === "'" && src[j + 1] !== "'")) j += src[j] === "'" ? 2 : 1;
      const m = /^!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/.exec(src.slice(j + 1));
      const text = src.slice(i, j + 1) + (m ? m[0] : '');
      out.push({ kind: 'ref', text });
      i = j + 1 + (m ? m[0].length : 0);
      continue;
    }
    if (ch === '[') {
      const j = src.indexOf(']', i);
      const end = j < 0 ? src.length : j + 1;
      const rest = /^[^!,()+\-*/^&=<>]*!?\$?[A-Za-z]{0,3}\$?\d*(?::\$?[A-Za-z]{1,3}\$?\d+)?/.exec(src.slice(end));
      out.push({ kind: 'ref', text: src.slice(i, end) + (rest ? rest[0] : '') });
      i = end + (rest ? rest[0].length : 0);
      continue;
    }
    const ref = /^(?:[A-Za-z_][\w.]*!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(])/.exec(src.slice(i));
    if (ref) {
      out.push({ kind: 'ref', text: ref[0] });
      i += ref[0].length;
      continue;
    }
    const err = /^#(REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)/i.exec(src.slice(i));
    if (err) {
      out.push({ kind: 'other', text: err[0].toUpperCase() });
      i += err[0].length;
      continue;
    }
    const num = /^\d+(?:\.\d+)?(?:[eE][-+]?\d+)?%?/.exec(src.slice(i));
    if (num) {
      out.push({ kind: 'num', text: num[0] });
      i += num[0].length;
      continue;
    }
    const word = /^[A-Za-z_][\w.]*/.exec(src.slice(i));
    if (word) {
      const isFunc = src[i + word[0].length] === '(';
      out.push({ kind: isFunc ? 'func' : 'other', text: word[0] });
      i += word[0].length;
      continue;
    }
    if (ch === '(') out.push({ kind: 'open', text: ch });
    else if (ch === ')') out.push({ kind: 'close', text: ch });
    else if (ch === ',' || ch === ';') out.push({ kind: 'comma', text: ch });
    else if (!/\s/.test(ch)) out.push({ kind: 'op', text: ch });
    i += 1;
  }
  return out;
}

const literalValue = (text: string): number => (text.endsWith('%') ? Number(text.slice(0, -1)) / 100 : Number(text));

export function analyzeFormula(formula: string): FormulaFinding {
  const tokens = tokenize(formula);
  const literals: LiteralHit[] = [];
  const externalWorkbooks = new Set<string>();
  const sheetRefs = new Set<string>();
  let hasRefError = false;
  // Stack of [functionName, argIndex] for each open paren.
  const stack: Array<{ fn: string; arg: number }> = [];
  let pendingFn = '';
  tokens.forEach((t, idx) => {
    if (t.kind === 'func') {
      pendingFn = t.text.toUpperCase();
      return;
    }
    if (t.kind === 'open') {
      stack.push({ fn: pendingFn, arg: 0 });
      pendingFn = '';
      return;
    }
    pendingFn = '';
    if (t.kind === 'close') {
      stack.pop();
      return;
    }
    if (t.kind === 'comma') {
      if (stack.length) stack[stack.length - 1].arg += 1;
      return;
    }
    if (t.kind === 'other' && t.text === '#REF!') hasRefError = true;
    if (t.kind === 'ref') {
      if (/#REF!/i.test(t.text)) hasRefError = true;
      const ext = /\[([^\]]+)\]/.exec(t.text);
      if (ext) externalWorkbooks.add(ext[1]);
      const sheet = /^'?(.+?)'?!/.exec(t.text.replace(/^\[[^\]]+\]/, ''));
      if (sheet && !ext) sheetRefs.add(sheet[1].replace(/''/g, "'"));
      return;
    }
    if (t.kind === 'num') {
      const value = literalValue(t.text);
      if (LITERAL_ALLOWLIST.has(value)) return;
      const frame = stack[stack.length - 1];
      if (frame && PARAMETER_ARGS[frame.fn]?.includes(frame.arg)) return;
      // A unary minus/plus before the number is part of it.
      const prev = tokens[idx - 1];
      const text = prev && prev.kind === 'op' && (prev.text === '-' || prev.text === '+') && (!tokens[idx - 2] || ['op', 'open', 'comma'].includes(tokens[idx - 2].kind)) ? prev.text + t.text : t.text;
      literals.push({ text, value: text.startsWith('-') ? -value : value });
    }
  });
  if (/#REF!/i.test(formula)) hasRefError = true;
  return { literals, hasRefError, externalWorkbooks: [...externalWorkbooks], sheetRefs: [...sheetRefs] };
}

/** Same formula pattern per column (R1C1-style) — the cell that breaks the pattern is the odd one. */
export function relativePattern(formula: string, row: number, col: number): string {
  const colIndex = (letters: string) => {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  };
  return formula.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])/g, (_m, dc, letters, dr, digits) => {
    const c = dc ? `C${colIndex(letters) + 1}` : `C[${colIndex(letters) - col}]`;
    const r = dr ? `R${digits}` : `R[${Number(digits) - 1 - row}]`;
    return r + c;
  });
}

export const ERROR_VALUES = /^#(REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!)$/;
