import React from 'react';

function renderBoldSegments(line: string): React.ReactNode {
  const segments = line.split(/\*\*/);
  return segments.map((seg, i) =>
    i % 2 === 1 ? (
      <strong key={i} className="cellix-md-strong">
        {seg}
      </strong>
    ) : (
      <React.Fragment key={i}>{seg}</React.Fragment>
    ),
  );
}

/** Single-line inline bold (`**term**`) — no block paragraph wrapper. */
export function renderInlineBoldMarkdown(text: string): React.ReactNode {
  const line = text.trim().split('\n')[0] ?? '';
  return renderBoldSegments(line);
}

/** Cursor-style: paragraphs + `**bold**`, single newlines → `<br />`. */
export function renderLightMarkdownPlain(text: string): React.ReactNode {
  const paragraphs = text.trim().split(/\n\n+/);
  return paragraphs.map((para, pi) => {
    const lines = para.split('\n');
    return (
      <p key={pi} className="cellix-response-md-p">
        {lines.map((line, li) => (
          <React.Fragment key={li}>
            {li > 0 && <br />}
            {renderBoldSegments(line)}
          </React.Fragment>
        ))}
      </p>
    );
  });
}
