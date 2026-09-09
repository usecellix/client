export function buildClientStatusMessage(message: string, sheetIsEmpty: boolean): string {
  const lower = message.toLowerCase();

  if (sheetIsEmpty) {
    if (/\b(create|generate|populate|add|fill|make|build|dummy|sample|row|header)\b/.test(lower)) {
      return 'Creating your table…';
    }
    return 'Working on your request…';
  }

  if (/\b(create|generate|populate)\b/.test(lower)) {
    return 'Updating your sheet…';
  }

  return 'Working on your request…';
}

export function isSimpleCreateTask(message: string, sheetIsEmpty: boolean): boolean {
  const lower = message.toLowerCase();
  return (
    sheetIsEmpty &&
    /\b(create|generate|add|populate|fill|make|build|give|dummy|sample)\b/.test(lower) &&
    /\b(row|rows|header|headers|column|columns|table)\b/.test(lower)
  );
}

/**
 * Fast, client-side mirror of the backend's `LlmRouterService.classifyIntent`
 * CHITCHAT category (greetings/small talk/thanks/goodbyes/identity questions)
 * — see `chitchat-prompt.ts`'s own definition, kept in sync with this list.
 * The backend already skips workbook context entirely for these and replies
 * in well under a second via plain `chunk` streaming, but the frontend had no
 * way to know that in advance, so it ran the SAME "Reading your worksheet…
 * Analyzing your spreadsheet structure…" choreography (with hard minimum
 * delays, `TIMING.readingMinRun`/`analyzingMinRun`, on top of the real wait)
 * for a bare "hi" as it does for a 20-sheet build. User report: "no need for
 * that... show thinking and show how do we handle it better."
 *
 * Deliberately conservative — a FALSE NEGATIVE here just means an ordinary
 * greeting gets the normal (slower) pacing, same as before this existed. A
 * FALSE POSITIVE would skip real analysis for an actual task, which is the
 * outcome to avoid, so this only matches short messages with no task-shaped
 * words at all (mirrors the backend prompt's own "TASK — anything referring
 * to the spreadsheet, data, formulas, formatting, or asking the assistant to
 * look at/analyze/change something" exclusion).
 */
const CHITCHAT_TASK_WORDS =
  /\b(sheet|workbook|cell|cells|row|rows|column|columns|table|data|formula|formulas|chart|graph|pivot|sum|total|average|count|sort|filter|create|generate|add|delete|remove|update|change|edit|fix|format|calculate|analyze|analyse|explain|summarize|summarise|find|search|compare|build|make|write|insert|export|import)\b/;

const CHITCHAT_GREETING_WORDS =
  /^(hi|hello|hey|yo|sup|hiya|greetings|good\s?morning|good\s?afternoon|good\s?evening|thanks|thank\s?you|thx|ty|bye|goodbye|see\s?ya|ok|okay|cool|nice|great|awesome|who\s+are\s+you|what\s+are\s+you|what\s+can\s+you\s+do|what\s+do\s+you\s+do)\b/i;

const MAX_CHITCHAT_LENGTH = 40;

export function isLikelyChitchat(message: string): boolean {
  const trimmed = message.trim();
  if (!trimmed || trimmed.length > MAX_CHITCHAT_LENGTH) return false;
  if (CHITCHAT_TASK_WORDS.test(trimmed)) return false;
  return CHITCHAT_GREETING_WORDS.test(trimmed) || /^[\p{Emoji}\s!?.]+$/u.test(trimmed);
}
