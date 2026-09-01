import React from 'react';

interface Span {
  start: number;
  end: number;
}

/**
 * Finds `=FUNCTION(...)`-shaped Excel formulas so they can render as inline
 * code even when the model didn't wrap them in backticks itself. Parens are
 * matched by depth rather than a non-greedy regex so nested calls like
 * `=IF(A1>0,SUM(B1:B10),0)` don't get cut off at the first `)`.
 */
function findFormulaSpans(line: string): Span[] {
  const spans: Span[] = [];
  const startRegex = /=[A-Za-z_][A-Za-z0-9_.]*\(/g;
  let match: RegExpExecArray | null;
  while ((match = startRegex.exec(line))) {
    const start = match.index;
    let i = start + match[0].length;
    let depth = 1;
    while (i < line.length && depth > 0) {
      if (line[i] === '(') depth++;
      else if (line[i] === ')') depth--;
      i++;
    }
    if (depth === 0) {
      spans.push({ start, end: i });
      startRegex.lastIndex = i;
    }
  }
  return spans;
}

type Token =
  | { type: 'text'; value: string }
  | { type: 'bold'; value: string }
  | { type: 'code'; value: string };

/** Tokenizes a single line into plain/bold/code runs: `` `code` ``, `**bold**`, and bare `=FORMULA(...)` spans. */
function tokenizeLine(line: string): Token[] {
  const tokens: Token[] = [];
  const formulaSpans = findFormulaSpans(line);
  let formulaIndex = 0;
  let i = 0;

  while (i < line.length) {
    if (line[i] === '`') {
      const close = line.indexOf('`', i + 1);
      if (close !== -1) {
        tokens.push({ type: 'code', value: line.slice(i + 1, close) });
        i = close + 1;
        continue;
      }
    }

    if (line.startsWith('**', i)) {
      const close = line.indexOf('**', i + 2);
      if (close !== -1) {
        tokens.push({ type: 'bold', value: line.slice(i + 2, close) });
        i = close + 2;
        continue;
      }
    }

    const span = formulaSpans[formulaIndex];
    if (span && span.start === i) {
      tokens.push({ type: 'code', value: line.slice(span.start, span.end) });
      i = span.end;
      formulaIndex++;
      continue;
    }

    let next = line.length;
    const backtickIdx = line.indexOf('`', i);
    if (backtickIdx !== -1) next = Math.min(next, backtickIdx);
    const boldIdx = line.indexOf('**', i);
    if (boldIdx !== -1) next = Math.min(next, boldIdx);
    if (span) next = Math.min(next, span.start);
    if (next <= i) next = i + 1;

    tokens.push({ type: 'text', value: line.slice(i, next) });
    i = next;
  }

  return tokens;
}

function renderTokens(tokens: Token[]): React.ReactNode {
  return tokens.map((token, i) => {
    if (token.type === 'code') {
      return (
        <code key={i} className="cellix-md-code">
          {token.value}
        </code>
      );
    }
    if (token.type === 'bold') {
      return (
        <strong key={i} className="cellix-md-strong">
          {token.value}
        </strong>
      );
    }
    return <React.Fragment key={i}>{token.value}</React.Fragment>;
  });
}

function renderLineSegments(line: string): React.ReactNode {
  return renderTokens(tokenizeLine(line));
}

/** Single-line inline formatting (`**bold**`, `` `code` ``, bare formulas) — no block paragraph wrapper. */
export function renderInlineBoldMarkdown(text: string): React.ReactNode {
  const line = text.trim().split('\n')[0] ?? '';
  return renderLineSegments(line);
}

/** Cursor-style: paragraphs + `**bold**` + inline code/formulas, single newlines → `<br />`. */
export function renderLightMarkdownPlain(text: string): React.ReactNode {
  const paragraphs = text.trim().split(/\n\n+/);
  return paragraphs.map((para, pi) => {
    const lines = para.split('\n');
    return (
      <p key={pi} className="cellix-response-md-p">
        {lines.map((line, li) => (
          <React.Fragment key={li}>
            {li > 0 && <br />}
            {renderLineSegments(line)}
          </React.Fragment>
        ))}
      </p>
    );
  });
}
