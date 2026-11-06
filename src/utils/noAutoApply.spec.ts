import { describe, expect, it } from 'vitest';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';

/**
 * P0.1 guard — "never auto-apply" must not depend on a flag.
 *
 * Model-proposed writes may only reach the sheet through the preview →
 * Accept path (previewManager.accept()). This test fails if the old
 * auto-apply flag or the write-then-revert preview helpers come back.
 */
const SRC_ROOT = resolve(__dirname, '..');
const SELF = resolve(__filename);

function walk(dir: string, out: string[] = []): string[] {
  for (const name of readdirSync(dir)) {
    const full = join(dir, name);
    if (statSync(full).isDirectory()) {
      walk(full, out);
    } else if (/\.(ts|tsx)$/.test(name) && resolve(full) !== SELF) {
      out.push(full);
    }
  }
  return out;
}

const files = walk(SRC_ROOT).map((file) => ({ file, text: readFileSync(file, 'utf8') }));

function offenders(pattern: RegExp): string[] {
  return files.filter(({ text }) => pattern.test(text)).map(({ file }) => file.replace(SRC_ROOT, 'src'));
}

describe('P0.1 — no auto-apply path', () => {
  it('has no autoApplyActions flag anywhere in client source', () => {
    expect(offenders(/autoApplyActions/)).toEqual([]);
  });

  it('has no shouldPreviewActions / previewPolicy gate left', () => {
    expect(offenders(/shouldPreviewActions|utils\/previewPolicy/)).toEqual([]);
  });

  it('has no write-then-revert preview helpers (previewActions / commitPreview / previewRanges)', () => {
    expect(offenders(/ActionEngine\.(previewActions|commitPreview|clearPreview)|previewRanges/)).toEqual([]);
  });

  it('does not apply actions straight from the conversation stream', () => {
    const stream = files.find(({ file }) => file.endsWith('hooks\\useConversation.ts') || file.endsWith('hooks/useConversation.ts'));
    expect(stream).toBeDefined();
    // The removed branch was `isActionMode && !usePreview && onActions`.
    expect(stream!.text).not.toMatch(/!usePreview/);
  });
});
