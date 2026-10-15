// @vitest-environment jsdom
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import { render } from '@testing-library/react';
import { renderInlineBoldMarkdown, renderLightMarkdownPlain, renderLightMarkdownWithRows } from './renderLightMarkdown';

describe('renderLightMarkdown', () => {
  it('wraps a bare =FUNCTION(...) formula in inline code', () => {
    const { container } = render(
      <>{renderInlineBoldMarkdown('Use =SUMIFS(C:C, A:A, "Ocean Polymers") for the total.')}</>,
    );
    const code = container.querySelector('code.cellix-md-code');
    expect(code?.textContent).toBe('=SUMIFS(C:C, A:A, "Ocean Polymers")');
  });

  it('keeps nested parens intact instead of cutting off at the first close-paren', () => {
    const { container } = render(<>{renderInlineBoldMarkdown('=IF(A1>0,SUM(B1:B10),0) is the formula.')}</>);
    const code = container.querySelector('code.cellix-md-code');
    expect(code?.textContent).toBe('=IF(A1>0,SUM(B1:B10),0)');
  });

  it('still renders explicit backtick code spans', () => {
    const { container } = render(<>{renderInlineBoldMarkdown('Set the value to `A1`.')}</>);
    const code = container.querySelector('code.cellix-md-code');
    expect(code?.textContent).toBe('A1');
  });

  it('still renders **bold** segments alongside formulas', () => {
    const { container } = render(
      <>{renderInlineBoldMarkdown('**Total:** =SUM(A1:A10)')}</>,
    );
    expect(container.querySelector('strong.cellix-md-strong')?.textContent).toBe('Total:');
    expect(container.querySelector('code.cellix-md-code')?.textContent).toBe('=SUM(A1:A10)');
  });

  it('does not treat plain "=" prose as a formula when not followed by a function call', () => {
    const { container } = render(<>{renderInlineBoldMarkdown('x = 5 is not a formula.')}</>);
    expect(container.querySelector('code.cellix-md-code')).toBeNull();
  });

  it('renders formulas across paragraph blocks via renderLightMarkdownPlain', () => {
    const { container } = render(
      <>{renderLightMarkdownPlain('First line.\n\nSecond paragraph with =COUNTA(A:A).')}</>,
    );
    expect(container.querySelectorAll('p.cellix-response-md-p')).toHaveLength(2);
    expect(container.querySelector('code.cellix-md-code')?.textContent).toBe('=COUNTA(A:A)');
  });
});

describe('renderLightMarkdown — bullets', () => {
  const html = (node: React.ReactNode) => renderToStaticMarkup(<div>{node}</div>);

  it('renders "• text" lines as two-column bullets so a wrapped bullet lines up under its first word', () => {
    const out = html(renderLightMarkdownPlain('Intro line.\n• First point that is long.\n• Second point.\nClosing line.'));
    expect(out.match(/cellix-md-bullet"/g)).toHaveLength(2);
    expect(out).toContain('cellix-md-bullet-mark');
    expect(out).toContain('cellix-md-bullet-text');
    // the mark is a column of its own, not characters glued to the text
    expect(out).not.toContain('• First');
  });

  it('adds no stray line break next to a bullet, but keeps breaks between ordinary lines', () => {
    const out = html(renderLightMarkdownPlain('A\nB\n• one\n• two\nC'));
    // only the A / B break remains; bullets are blocks, so none is needed around them
    expect((out.match(/<br\/?>/g) ?? []).length).toBe(1);
  });

  it('a half-typed bullet mark stays plain text until the space after it arrives', () => {
    expect(html(renderLightMarkdownPlain('•'))).not.toContain('cellix-md-bullet');
    expect(html(renderLightMarkdownPlain('• x'))).toContain('cellix-md-bullet');
  });

  it('works in the variant that turns "(row N)" into pointers, too', () => {
    const out = html(renderLightMarkdownWithRows('• see (row 7) now', (n, key) => <b key={key}>R{n}</b>));
    expect(out).toContain('cellix-md-bullet');
    expect(out).toContain('<b>R7</b>');
  });
});
