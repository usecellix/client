import { SheetAction } from '@/hooks/useSseStream';

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
  | { type: 'actions'; data: SseActionsData }
  | { type: 'error'; data: { message: string } }
  | { type: 'conversation_end'; data: { summary?: string; conversationId?: string } }
  | { type: 'done'; data: { message: string } };

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

  const normalizedType = (eventType || '').trim();
  const inferredType =
    normalizedType === ''
      ? parsed && Array.isArray(parsed.actions)
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
    case 'status':
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
