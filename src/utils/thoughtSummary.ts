export type ThoughtPhase = 'reading' | 'analyzing' | 'composing' | 'streaming' | 'final';

function matchesAny(lower: string, terms: string[]): boolean {
  return terms.some((t) => lower.includes(t));
}

export function buildThoughtSummary(
  userMessage: string,
  phase: ThoughtPhase,
): string {
  const lower = userMessage.trim().toLowerCase();

  if (phase === 'reading') {
    if (matchesAny(lower, ['create', 'generate', 'populate', 'dummy', 'sample', 'header', 'row'])) {
      return 'Preparing your table…';
    }
    return 'Reading the active worksheet…';
  }

  if (phase === 'analyzing') {
    if (matchesAny(lower, ['create', 'generate', 'populate', 'dummy', 'sample', 'header', 'row'])) {
      return 'Building headers and rows…';
    }
    if (matchesAny(lower, ['explain', 'describe', 'overview', 'summarize'])) {
      return [
        'Analyzing your spreadsheet structure and cell values.',
        'Identifying headers, data density, and patterns worth calling out in plain language…',
      ].join('\n\n');
    }
    if (matchesAny(lower, ['total', 'sum', 'calculate', 'add up'])) {
      return [
        'Looking for numeric columns and valid totals.',
        'Checking which column you mean before writing anything to the sheet…',
      ].join('\n\n');
    }
    if (matchesAny(lower, ['sort', 'filter', 'find', 'match'])) {
      return [
        'Scanning rows for matches and comparing values across columns.',
        'Planning how to present or apply the results on your sheet…',
      ].join('\n\n');
    }
    return [
      'Analyzing your spreadsheet structure and understanding the data.',
      'Deciding the clearest way to answer your request…',
    ].join('\n\n');
  }

  if (phase === 'composing') {
    return [
      'Good. I have enough context from the sheet.',
      'Composing a clear response you can review before any changes are applied…',
    ].join('\n\n');
  }

  if (phase === 'streaming') {
    return [
      'Working through your request.',
      'Drafting the answer and double-checking it against the sheet data…',
    ].join('\n\n');
  }

  // final
  if (matchesAny(lower, ['explain', 'describe', 'overview', 'summarize'])) {
    return [
      'Good. I finished reviewing your sheet layout and cell contents.',
      'Summarized what the data represents and highlighted useful next steps.',
      'Prepared the explanation below for you to review.',
    ].join('\n\n');
  }
  if (matchesAny(lower, ['total', 'sum', 'calculate'])) {
    return [
      'Good. I identified the relevant numeric column and checked the values.',
      'Calculated the total and prepared the result below.',
    ].join('\n\n');
  }
  if (matchesAny(lower, ['sort', 'filter', 'match'])) {
    return [
      'Good. I reviewed the matching rows and column structure.',
      'Prepared the results and next steps below.',
    ].join('\n\n');
  }

  return [
    'Good. I finished reviewing your worksheet.',
    'Prepared the response below for you to review.',
  ].join('\n\n');
}
