import { getToolResultEndpoint } from '@/lib/apiConfig';
import { fetchRangeData } from '@/services/rangeFetchService';

export interface ToolRequestPayload {
  requestId: string;
  conversationId: string;
  tool: string;
  sheet: string;
  range: string;
}

export async function handleToolRequest(payload: ToolRequestPayload): Promise<void> {
  if (payload.tool !== 'get_range_data') return;

  const endpoint = getToolResultEndpoint();
  const headers: Record<string, string> = { 'Content-Type': 'application/json' };
  if (endpoint.includes('.ngrok-free.app')) {
    headers['ngrok-skip-browser-warning'] = 'true';
  }

  try {
    const { values } = await fetchRangeData(payload.sheet, payload.range);
    await fetch(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        conversationId: payload.conversationId,
        requestId: payload.requestId,
        tool: payload.tool,
        values,
      }),
    });
  } catch (error) {
    await fetch(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        conversationId: payload.conversationId,
        requestId: payload.requestId,
        tool: payload.tool,
        error: error instanceof Error ? error.message : 'Range fetch failed',
      }),
    });
  }
}
