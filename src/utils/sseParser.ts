import { SheetAction } from '@/hooks/useSseStream';
import { CellChange } from '@/types/changeSet';
import { ClarificationPayload } from '@/types/cellix.types';

export interface SseClarificationData extends ClarificationPayload {
  conversationId?: string;
}

export interface SseAnswerData {
  answer: string;
  matches?: SseMatchData[];
  conversationId?: string;
}

export interface SseQuestionData {
  question: string;
  options?: string[];
  conversationId?: string;
}

export interface SseActionsData {
  actions: SheetAction[];
  explanation: string;
  conversationId?: string;
  changeSetId?: string;
  changes?: CellChange[];
}

export interface SseToolRequestData {
  requestId: string;
  conversationId: string;
  tool: string;
  sheet: string;
  range: string;
}

export interface SsePlanStepData {
  title: string;
  detail?: string;
}

export interface SsePlanData {
  prompt: string;
  summary?: string;
  steps: SsePlanStepData[];
  affectedSheets: string[];
  estimatedRows?: number;
  safestApproach?: string;
  conversationId?: string;
}

export interface SseMatchData {
  label: string;
  detail?: string;
  sheetName: string;
  row: number;
  col: number;
  colLetter?: string;
  rowNum?: number;
}

export interface SseMatchesData {
  matches: SseMatchData[];
  summary?: string;
  conversationId?: string;
}

export interface SseSelectCellData {
  sheetName: string;
  row: number;
  col: number;
  conversationId?: string;
}

export type ParsedSseEvent =
  | { type: 'status'; data: { message: string } }
  | { type: 'chunk'; data: { text: string } }
  | { type: 'answer'; data: SseAnswerData }
  | { type: 'question'; data: SseQuestionData }
  | { type: 'clarification'; data: SseClarificationData }
  | { type: 'actions'; data: SseActionsData }
  | { type: 'tool_request'; data: SseToolRequestData }
  | { type: 'plan'; data: SsePlanData }
  | { type: 'matches'; data: SseMatchesData }
  | { type: 'select_cell'; data: SseSelectCellData }
  | { type: 'error'; data: { message: string } }
  | { type: 'conversation_end'; data: { summary?: string; conversationId?: string } }
  | { type: 'done'; data: { message: string } };

function normalizeClarificationPayload(
  payload: Record<string, unknown>,
): SseClarificationData | null {
  const question = String(payload.question ?? '');
  if (!question) return null;

  return {
    question,
    suggestions: Array.isArray(payload.suggestions)
      ? (payload.suggestions as string[])
      : undefined,
    ambiguityScore: Number(payload.ambiguityScore ?? 0),
    conversationId: payload.conversationId as string | undefined,
  };
}

function parseMasterEnvelope(parsed: Record<string, unknown>): ParsedSseEvent | null {
  const innerType = String(parsed.type ?? '');
  const payload = parsed.payload;
  if (!innerType || !payload || typeof payload !== 'object') return null;

  const p = payload as Record<string, unknown>;

  switch (innerType) {
    case 'clarification': {
      const data = normalizeClarificationPayload(p);
      return data ? { type: 'clarification', data } : null;
    }
    case 'thinking':
      return {
        type: 'status',
        data: { message: String(p.message ?? 'Thinking…') },
      };
    case 'actions':
      return Array.isArray(p.actions)
        ? {
            type: 'actions',
            data: {
              actions: p.actions as SheetAction[],
              explanation: String(p.summary ?? p.explanation ?? ''),
              conversationId: p.conversationId as string | undefined,
              changeSetId: p.changeSetId as string | undefined,
              changes: Array.isArray(p.changes) ? (p.changes as CellChange[]) : undefined,
            },
          }
        : null;
    case 'error':
      return {
        type: 'error',
        data: { message: String(p.message ?? 'Unknown error') },
      };
    case 'done':
      return { type: 'done', data: { message: 'done' } };
    default:
      return null;
  }
}

export function parseSseEventBlock(block: string): ParsedSseEvent | null {
  const lines = block.split(/\r?\n/);
  let eventType = '';
  const dataLines: string[] = [];

  for (const line of lines) {
    if (line.startsWith('event:')) {
      eventType = line.slice('event:'.length).trim();
    } else if (line.startsWith('data:')) {
      dataLines.push(line.slice('data:'.length).trimStart());
    }
  }

  const rawData = dataLines.join('\n');
  if (!eventType && !rawData) return null;

  let parsed: Record<string, unknown> | null = null;
  try {
    parsed = rawData ? (JSON.parse(rawData) as Record<string, unknown>) : null;
  } catch {
    parsed = null;
  }

  if (parsed && typeof parsed.type === 'string' && parsed.payload) {
    const envelope = parseMasterEnvelope(parsed);
    if (envelope) return envelope;
  }

  const normalizedType = (eventType || '').trim();
  const inferredType =
    normalizedType === ''
      ? parsed && typeof parsed.type === 'string' && parsed.type === 'clarification'
        ? 'clarification'
        : parsed && Array.isArray(parsed.actions)
          ? 'actions'
          : 'chunk'
      : normalizedType;

  switch (inferredType) {
    case 'chunk':
      return { type: 'chunk', data: (parsed as { text: string }) ?? { text: rawData } };
    case 'answer':
      return {
        type: 'answer',
        data: parsed
          ? ({
              answer: String(parsed.answer ?? rawData),
              matches: Array.isArray(parsed.matches)
                ? (parsed.matches as SseMatchData[])
                : undefined,
              conversationId: parsed.conversationId as string | undefined,
            })
          : { answer: rawData },
      };
    case 'question':
      return parsed
        ? {
            type: 'question',
            data: {
              question: String(parsed.question ?? rawData),
              options: Array.isArray(parsed.options) ? (parsed.options as string[]) : undefined,
              conversationId: parsed.conversationId as string | undefined,
            },
          }
        : { type: 'question', data: { question: rawData } };
    case 'clarification':
      if (parsed) {
        const data = normalizeClarificationPayload(parsed);
        if (data) return { type: 'clarification', data };
      }
      return null;
    case 'status':
      return {
        type: 'status',
        data: parsed ? { message: String(parsed.message ?? rawData) } : { message: rawData },
      };
    case 'thinking':
      return {
        type: 'status',
        data: parsed ? { message: String(parsed.message ?? rawData) } : { message: rawData },
      };
    case 'actions':
      return parsed && Array.isArray(parsed.actions)
        ? {
            type: 'actions',
            data: {
              actions: parsed.actions as SheetAction[],
              explanation: String(parsed.explanation ?? ''),
              conversationId: parsed.conversationId as string | undefined,
              changeSetId: parsed.changeSetId as string | undefined,
              changes: Array.isArray(parsed.changes)
                ? (parsed.changes as CellChange[])
                : undefined,
            },
          }
        : null;
    case 'tool_request':
      return parsed &&
        typeof parsed.requestId === 'string' &&
        typeof parsed.sheet === 'string' &&
        typeof parsed.range === 'string'
        ? {
            type: 'tool_request',
            data: {
              requestId: String(parsed.requestId),
              conversationId: String(parsed.conversationId ?? ''),
              tool: String(parsed.tool ?? 'get_range_data'),
              sheet: String(parsed.sheet),
              range: String(parsed.range),
            },
          }
        : null;
    case 'plan':
      return parsed && Array.isArray(parsed.steps)
        ? {
            type: 'plan',
            data: {
              prompt: String(parsed.prompt ?? ''),
              summary: parsed.summary ? String(parsed.summary) : undefined,
              steps: (parsed.steps as SsePlanStepData[]).map((s) => ({
                title: String(s.title ?? ''),
                detail: s.detail ? String(s.detail) : undefined,
              })),
              affectedSheets: Array.isArray(parsed.affectedSheets)
                ? (parsed.affectedSheets as string[]).map(String)
                : [],
              estimatedRows:
                typeof parsed.estimatedRows === 'number' ? parsed.estimatedRows : undefined,
              safestApproach: parsed.safestApproach
                ? String(parsed.safestApproach)
                : undefined,
              conversationId: parsed.conversationId as string | undefined,
            },
          }
        : null;
    case 'matches':
      return parsed && Array.isArray(parsed.matches)
        ? {
            type: 'matches',
            data: {
              matches: (parsed.matches as SseMatchData[]).map((m) => ({
                label: String(m.label ?? ''),
                detail: m.detail ? String(m.detail) : undefined,
                sheetName: String(m.sheetName ?? ''),
                row: Number(m.row ?? 0),
                col: Number(m.col ?? 0),
                colLetter: m.colLetter ? String(m.colLetter) : undefined,
                rowNum: typeof m.rowNum === 'number' ? m.rowNum : undefined,
              })),
              summary: parsed.summary ? String(parsed.summary) : undefined,
              conversationId: parsed.conversationId as string | undefined,
            },
          }
        : null;
    case 'select_cell':
      return parsed
        ? {
            type: 'select_cell',
            data: {
              sheetName: String(parsed.sheetName ?? ''),
              row: Number(parsed.row ?? 0),
              col: Number(parsed.col ?? 0),
              conversationId: parsed.conversationId as string | undefined,
            },
          }
        : null;
    case 'error':
      return { type: 'error', data: (parsed as { message: string }) ?? { message: rawData || 'Unknown error' } };
    case 'conversation_end':
      return {
        type: 'conversation_end',
        data: (parsed as { summary?: string }) ?? { summary: rawData || 'done' },
      };
    case 'done':
      return { type: 'done', data: (parsed as { message: string }) ?? { message: rawData || 'done' } };
    default:
      return null;
  }
}
