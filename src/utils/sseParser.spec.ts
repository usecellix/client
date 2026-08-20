import { describe, expect, it } from 'vitest';
import { parseSseEventBlock } from '@/utils/sseParser';

describe('parseSseEventBlock — actions event', () => {
  it('carries irreversibleActionTypes through from the raw SSE payload', () => {
    const block =
      'event: actions\n' +
      `data: ${JSON.stringify({
        actions: [{ type: 'RENAME_SHEET', sheetName: 'Sheet1', newName: 'Renamed' }],
        explanation: 'Renamed the sheet.',
        changeSetId: 'cs_1',
        irreversibleActionTypes: ['RENAME_SHEET'],
      })}`;

    const event = parseSseEventBlock(block);

    expect(event?.type).toBe('actions');
    expect(event?.type === 'actions' && event.data.irreversibleActionTypes).toEqual([
      'RENAME_SHEET',
    ]);
  });

  it('leaves irreversibleActionTypes undefined when the backend omits it', () => {
    const block =
      'event: actions\n' +
      `data: ${JSON.stringify({
        actions: [{ type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 5 }],
        explanation: 'Set A1 to 5.',
        changeSetId: 'cs_2',
      })}`;

    const event = parseSseEventBlock(block);

    expect(event?.type).toBe('actions');
    expect(event?.type === 'actions' && event.data.irreversibleActionTypes).toBeUndefined();
  });
});
