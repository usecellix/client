import { useState, useCallback, useRef } from 'react';
import { getConversationEndpoint } from '@/lib/apiConfig';
import { prepareConversationRequestPayload } from '@/utils/payloadCompressor';

// SSE event types as per PRD
export interface StatusEvent {
  message: string;
}

export interface ChunkEvent {
  text: string;
}

export interface ActionsEvent {
  actions: SheetAction[];
  explanation: string;
}

export interface ErrorEvent {
  message: string;
}

export interface DoneEvent {
  message: string;
}

export interface AnswerEvent {
  answer: string;
  conversationId?: string;
}

export interface QuestionEvent {
  question: string;
  options?: string[];
  conversationId?: string;
}

export interface ConversationEndEvent {
  summary?: string;
  conversationId?: string;
}

import { SheetAction } from '@/types/sheet-actions';

export type { SheetAction };

// Union type for all SSE events
export type SseEvent = 
  | { type: 'status'; data: StatusEvent }
  | { type: 'chunk'; data: ChunkEvent }
  | { type: 'answer'; data: AnswerEvent }
  | { type: 'question'; data: QuestionEvent }
  | { type: 'actions'; data: ActionsEvent }
  | { type: 'error'; data: ErrorEvent }
  | { type: 'done'; data: DoneEvent }
  | { type: 'conversation_end'; data: ConversationEndEvent };

// Hook state interface
interface UseSseStreamState {
  isConnected: boolean;
  statusMessages: string[];
  aiText: string;
  actions: SheetAction[];
  explanation: string;
  error: string | null;
  isComplete: boolean;
}

// Hook return interface
interface UseSseStreamReturn extends UseSseStreamState {
  startStream: (prompt: string, sheetData: any[][]) => Promise<void>;
  stopStream: () => void;
  reset: () => void;
}

export const useSseStream = (): UseSseStreamReturn => {
  const [state, setState] = useState<UseSseStreamState>({
    isConnected: false,
    statusMessages: [],
    aiText: '',
    actions: [],
    explanation: '',
    error: null,
    isComplete: false,
  });

  const abortControllerRef = useRef<AbortController | null>(null);

  const parseSseEventBlock = useCallback((block: string): SseEvent | null => {
    // SSE block format: lines like `event: chunk`, `data: {...}` (data may repeat)
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

    const tryParseJson = (): any => {
      try {
        return JSON.parse(rawData);
      } catch {
        return null;
      }
    };

    const parsed = tryParseJson();
    const normalizedType = (eventType || '').trim();

    // Backend Phase 0 sometimes emits a blank/space event name for the actions payload.
    const inferredType =
      normalizedType === ''
        ? parsed && typeof parsed === 'object' && Array.isArray(parsed.actions)
          ? 'actions'
          : 'chunk'
        : normalizedType;

    switch (inferredType) {
      case 'chunk':
        return { type: 'chunk', data: parsed ?? { text: rawData } };
      case 'answer':
        return { type: 'answer', data: parsed ?? { answer: rawData } };
      case 'question':
        return parsed ? { type: 'question', data: parsed } : { type: 'question', data: { question: rawData } };
      case 'status':
        return { type: 'status', data: parsed ?? { message: rawData } };
      case 'actions':
        return parsed ? { type: 'actions', data: parsed } : null;
      case 'error':
        return { type: 'error', data: parsed ?? { message: rawData || 'Unknown error' } };
      case 'done':
        return { type: 'done', data: parsed ?? { message: rawData || 'done' } };
      case 'conversation_end':
        return { type: 'conversation_end', data: parsed ?? { summary: rawData || 'done' } };
      default:
        return parsed ? ({ type: inferredType as any, data: parsed } as SseEvent) : null;
    }
  }, []);

  const getUserFacingErrorMessage = useCallback(async (response: Response): Promise<string> => {
    const traceIdHeader = response.headers.get('x-trace-id');
    const contentType = response.headers.get('content-type') || '';

    const readJsonSafely = async (): Promise<any | null> => {
      try {
        return await response.clone().json();
      } catch {
        return null;
      }
    };

    const json = contentType.includes('application/json') ? await readJsonSafely() : null;
    const maybeTraceId = (json && typeof json.traceId === 'string' && json.traceId) || traceIdHeader;

    const isErrorEnvelope =
      json &&
      json.success === false &&
      typeof json.traceId === 'string' &&
      json.error &&
      typeof json.error.message === 'string';

    if (isErrorEnvelope) {
      // eslint-disable-next-line no-console
      console.error(`[Cellix] Backend error traceId=${json.traceId}`, json.error);
      return json.error.message as string;
    }

    if (maybeTraceId) {
      // eslint-disable-next-line no-console
      console.error(`[Cellix] Request failed traceId=${maybeTraceId}`);
    }

    if (json && typeof json.message === 'string') return json.message;

    let text = '';
    try {
      text = await response.text();
    } catch {
      text = '';
    }
    return text?.trim() ? text.trim() : `Request failed (HTTP ${response.status})`;
  }, []);

  // Handle individual SSE events
  const handleEvent = useCallback((event: SseEvent) => {
    setState(prevState => {
      switch (event.type) {
        case 'status':
          return {
            ...prevState,
            statusMessages: [...prevState.statusMessages, event.data.message],
          };
          
        case 'chunk':
          return {
            ...prevState,
            aiText: prevState.aiText + event.data.text,
          };

        case 'answer':
          return {
            ...prevState,
            aiText: prevState.aiText + event.data.answer,
          };

        case 'question': {
          const optionsText = event.data.options?.length ? `\n\n${event.data.options.join('\n')}` : '';
          return {
            ...prevState,
            aiText: `${prevState.aiText}${event.data.question}${optionsText}`,
            isConnected: false,
            isComplete: true,
          };
        }
          
        case 'actions':
          return {
            ...prevState,
            actions: event.data.actions,
            explanation: event.data.explanation,
          };
          
        case 'error':
          return {
            ...prevState,
            error: event.data.message,
            isConnected: false,
            isComplete: true,
          };
          
        case 'done':
        case 'conversation_end':
          return {
            ...prevState,
            isConnected: false,
            isComplete: true,
          };
          
        default:
          return prevState;
      }
    });
  }, []);

  // Start SSE stream
  const startStream = useCallback(async (prompt: string, sheetData: any[][]) => {
    // Reset state
    setState({
      isConnected: true,
      statusMessages: [],
      aiText: '',
      actions: [],
      explanation: '',
      error: null,
      isComplete: false,
    });

    // Create new AbortController
    abortControllerRef.current = new AbortController();

    try {
      const requestPayload = prepareConversationRequestPayload(prompt, sheetData);
      const streamEndpoint = getConversationEndpoint();
      const headers: Record<string, string> = {
        'Content-Type': 'application/json',
      };

      if (streamEndpoint.includes('.ngrok-free.app')) {
        headers['ngrok-skip-browser-warning'] = 'true';
      }

      const response = await fetch(streamEndpoint, {
        method: 'POST',
        headers,
        credentials: 'include',
        body: JSON.stringify(requestPayload),
        signal: abortControllerRef.current.signal,
      });

      if (!response.ok) {
        const message = await getUserFacingErrorMessage(response);
        throw new Error(message);
      }

      const reader = response.body?.getReader();
      const decoder = new TextDecoder();
      let buffer = '';

      if (!reader) {
        throw new Error('Response body is not readable');
      }

      // Read stream
      while (true) {
        const { done, value } = await reader.read();
        
        if (done) break;

        // Decode chunk and add to buffer
        buffer += decoder.decode(value, { stream: true });

        // Process complete SSE events separated by blank line
        const parts = buffer.split(/\r?\n\r?\n/);
        buffer = parts.pop() || '';
        for (const part of parts) {
          const evt = parseSseEventBlock(part);
          if (evt) handleEvent(evt);
        }
      }

      // Handle any remaining data in buffer
      if (buffer.trim()) {
        const evt = parseSseEventBlock(buffer);
        if (evt) handleEvent(evt);
      }

      // End-of-stream: finalize even if backend doesn't emit `done`
      handleEvent({ type: 'done', data: { message: 'done' } });

    } catch (error: any) {
      if (error.name === 'AbortError') {
        console.log('Stream was aborted');
        setState(prevState => ({
          ...prevState,
          isConnected: false,
          isComplete: true,
        }));
      } else {
        const endpoint = getConversationEndpoint();
        const message =
          error.message === 'Failed to fetch'
            ? `Failed to fetch ${endpoint}. Make sure the Vite dev server was restarted and its /api proxy can reach the backend.`
            : error.message || 'Stream connection failed';

        console.error('Stream error:', {
          endpoint,
          pageOrigin: window.location.origin,
          error,
        });
        setState(prevState => ({
          ...prevState,
          error: message,
          isConnected: false,
          isComplete: true,
        }));
      }
    } finally {
      abortControllerRef.current = null;
    }
  }, [parseSseEventBlock, handleEvent, getUserFacingErrorMessage]);

  // Stop active stream
  const stopStream = useCallback(() => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }
  }, []);

  // Reset state
  const reset = useCallback(() => {
    stopStream();
    setState({
      isConnected: false,
      statusMessages: [],
      aiText: '',
      actions: [],
      explanation: '',
      error: null,
      isComplete: false,
    });
  }, [stopStream]);

  return {
    ...state,
    startStream,
    stopStream,
    reset,
  };
};
