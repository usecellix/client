import React from 'react';

export type FileBadgeKind = 'pdf' | 'sheet';

interface FileTypeBadgeProps {
  kind: FileBadgeKind;
  size?: number;
}

/**
 * A document icon with the file type on a coloured banner: red PDF, green
 * sheet. Drawn as SVG so it stays
 * sharp at any size and sits on any background.
 */
const FileTypeBadge: React.FC<FileTypeBadgeProps> = ({ kind, size = 64 }) => {
  const banner = kind === 'pdf' ? '#e53935' : '#1e8e5a';
  const label = kind === 'pdf' ? 'PDF' : 'XLS';
  return (
    <svg
      className={`cellix-file-badge is-${kind}`}
      width={size}
      height={size}
      viewBox="0 0 64 64"
      role="img"
      aria-label={label}
    >
      <path
        d="M14 4h25l13 13v39a4 4 0 0 1-4 4H14a4 4 0 0 1-4-4V8a4 4 0 0 1 4-4z"
        fill="#ffffff"
        stroke="#cbd5e1"
        strokeWidth="1.5"
      />
      <path d="M39 4v9a4 4 0 0 0 4 4h9z" fill="#e2e8f0" />
      <g stroke="#cbd5e1" strokeWidth="1.5" strokeLinecap="round">
        <path d="M17 17h16M17 22h30M17 27h30" />
      </g>
      <rect x="4" y="33" width="46" height="17" rx="2.5" fill={banner} />
      <text
        x="27"
        y="45.5"
        textAnchor="middle"
        fontSize="12.5"
        fontWeight="700"
        fontFamily="Segoe UI, Arial, sans-serif"
        fill="#ffffff"
        letterSpacing="0.5"
      >
        {label}
      </text>
    </svg>
  );
};

export default FileTypeBadge;
