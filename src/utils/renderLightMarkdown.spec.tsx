// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { render } from '@testing-library/react';
import { renderInlineBoldMarkdown, renderLightMarkdownPlain } from './renderLightMarkdown';

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
