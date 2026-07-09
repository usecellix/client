/** Sheet tags inserted by the conversation UI, e.g. @[Summary] */
const SHEET_MENTION_PATTERN = /@\[([^\]]+)\]/g;

export function extractSheetMentions(message: string): string[] {
  const names: string[] = [];
  for (const match of message.matchAll(SHEET_MENTION_PATTERN)) {
    const name = match[1]?.trim();
    if (name && !names.includes(name)) {
      names.push(name);
    }
  }
  return names;
}

/** Remove @[Sheet] tags; optional quoted replacement keeps the name for LLM routing. */
export function stripSheetMentions(message: string, replacement: '' | 'name' = 'name'): string {
  if (replacement === '') {
    return message.replace(SHEET_MENTION_PATTERN, '').replace(/\s{2,}/g, ' ').trim();
  }
  return message.replace(SHEET_MENTION_PATTERN, '$1').replace(/\s{2,}/g, ' ').trim();
}
