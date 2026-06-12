import { SheetAction } from '@/hooks/useSseStream';
import { ClarificationPayload } from '@/types/cellix.types';

export interface SseClarificationData extends ClarificationPayload {
  conversationId?: string;
}

export interface SseAnswerData {
  answer: string;
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
}

export type ParsedSseEvent =
  | { type: 'status'; data: { message: string } }
  | { type: 'chunk'; data: { text: string } }
  | { type: 'answer'; data: SseAnswerData }
  | { type: 'question'; data: SseQuestionData }
  | { type: 'clarification'; data: SseClarificationData }
  | { type: 'actions'; data: SseActionsData }
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
          ? ({ answer: String(parsed.answer ?? rawData), conversationId: parsed.conversationId as string | undefined })
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
