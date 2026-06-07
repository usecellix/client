const DEFAULT_FOLLOW_UPS = [
  'Explain this sheet to me',
  'Calculate the total for a column',
  'Highlight cells that look unusual',
];

export function generateSuggestedFollowUps(answer: string, userPrompt?: string): string[] {
  const text = `${answer} ${userPrompt ?? ''}`.toLowerCase();
  const suggestions: string[] = [];

  const columnMatch = answer.match(/\bcolumn\s+([A-Za-z0-9 ]+)/i) ?? userPrompt?.match(/\bcolumn\s+([A-Za-z0-9 ]+)/i);
  const columnLabel = columnMatch?.[1]?.trim().split(/\s+/)[0]?.toUpperCase();

  if (/which column|what column|pick a column|choose a column/i.test(answer)) {
    if (columnLabel) {
      suggestions.push(`Calculate the total for column ${columnLabel}`);
      suggestions.push(`Sort the sheet by column ${columnLabel} values`);
      suggestions.push(`Clear empty cells in column ${columnLabel}`);
    } else {
      suggestions.push('Use column A for the calculation');
      suggestions.push('Use column B for the calculation');
      suggestions.push('Explain what each column contains');
    }
    return suggestions.slice(0, 3);
  }

  if (columnLabel) {
    suggestions.push(`Sort the sheet by column ${columnLabel} values`);
    suggestions.push(`Clear all empty cells and consolidate data in column ${columnLabel}`);
    suggestions.push(`Delete all rows that are completely empty`);
  }

  if (/total|sum|calculate/i.test(text)) {
    suggestions.push('Show the total in a summary cell');
    suggestions.push('Break down totals by column');
  }

  if (/cell\s+[A-Za-z]+\d+/i.test(answer)) {
    suggestions.push('What other cells should I check?');
    suggestions.push('Summarize this column');
  }

  if (/empty|blank/i.test(text)) {
    suggestions.push('Delete all rows that are completely empty');
    suggestions.push('Fill blank cells with a default value');
  }

  if (suggestions.length === 0) {
    return DEFAULT_FOLLOW_UPS;
  }

  const unique = Array.from(new Set(suggestions));
  while (unique.length < 3) {
    const fallback = DEFAULT_FOLLOW_UPS.find((item) => !unique.includes(item));
    if (!fallback) break;
    unique.push(fallback);
  }

  return unique.slice(0, 3);
}
