export type ThoughtPhase = 'reading' | 'analyzing' | 'composing' | 'streaming' | 'final';

/**
 * Whether a turn's "Thought process" content is important enough to auto-open
 * rather than sit collapsed. This is the ONLY thing standing between a
 * build-incompleteness warning and a user who never expands the disclosure —
 * `pushHistory`/`updateTurn` route these warnings into the thinking log
 * (never hidden by `hideProgress` once it has real content — TurnRenderer's
 * own rule), but a COLLAPSED thinking block is functionally invisible to
 * anyone who doesn't think to click it.
 *
 * Live incident (Sept 9, 2026): a stepwise build's Planner response truncated
 * (again — the same root cause TASKS.md #187 pruning was built for), pruning
 * 7 of 10 subtasks down to 3 to avoid shipping broken formulas. The backend
 * correctly said so — "I could not fully plan 7 steps of this request... ask
 * again to add the missing part" — but that text only ever reached the
 * collapsed thinking log, whose OLD auto-expand check
 * (`/blocked|cannot create|verification/i`) doesn't recognize this wording at
 * all. The user saw two "Applied" cards and nothing telling them 70% of their
 * request never got built. Widened here to cover every existing
 * completeness-warning phrasing this codebase already emits (TASKS.md #171's
 * "proceeded under an assumption", #155's "produced no changes", #187's
 * "could not fully plan").
 */
export function isCompletenessWarningThought(content: string): boolean {
  return /blocked|cannot create|verification|could not fully plan|produced no changes|proceeding under an assumption|things? to confirm/i.test(
    content,
  );
}

/**
 * Escalating "still working" copy for a long, opaque wait — a single
 * non-streaming Planner call for a large build (12+ month sheets, a
 * dashboard) can take 30 seconds to several minutes with LITERALLY nothing
 * for the backend to report in between (it is one request/response, not a
 * stream), so the existing status line otherwise sits unchanged the whole
 * time. A live report (Sept 8, 2026): "user is sitting idle... no loading or
 * anything."
 *
 * Deliberately generic and honest rather than fake-specific — this never
 * claims to know what step the backend is on, only that a longer wait is
 * normal for a bigger ask. A real backend status/thinking event always
 * supersedes this (see `TurnRuntime.lastLiveUpdateAt` in useConversation.ts)
 * — the ticker only fills the gaps between real updates, never overwrites one
 * that just arrived.
 */
const STILL_WORKING_TIERS: Array<{ afterMs: number; message: string }> = [
  { afterMs: 15_000, message: 'Still working — breaking this into steps for a build this size…' },
  {
    afterMs: 40_000,
    message: 'This looks like a multi-part build — it can take a couple of minutes…',
  },
  {
    afterMs: 90_000,
    message:
      "Still going — large builds like this sometimes take a few minutes. Feel free to keep this tab open; the first step will show as soon as it's ready…",
  },
];

/**
 * The escalating message for `elapsedMs` of silence, or `null` before the
 * first tier's threshold (the existing scripted reveal already covers early
 * seconds — this only kicks in once a wait has gone on long enough to risk
 * feeling frozen).
 */
export function stillWorkingMessage(elapsedMs: number): string | null {
  let current: string | null = null;
  for (const tier of STILL_WORKING_TIERS) {
    if (elapsedMs >= tier.afterMs) current = tier.message;
  }
  return current;
}

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
