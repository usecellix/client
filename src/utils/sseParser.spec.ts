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

/**
 * TASKS.md #393. The server sends an answer's row pointers in the `answer`
 * event and again in a `matches` event, and the second replaces the first in
 * the add-in. This parser rebuilt each pointer without `endCol`, so a click on
 * "row 445" selected A445 instead of A445:I445.
 */
describe('parseSseEventBlock — matches event', () => {
  const rowPointer = {
    label: 'row 445',
    sheetName: 'Bank Statement',
    row: 444,
    col: 0,
    colLetter: 'A',
    rowNum: 445,
    rawValue: '',
    endCol: 8,
    detail: 'A445:I445',
  };

  it('keeps endCol on a pointer that covers a whole row', () => {
    const event = parseSseEventBlock(
      `event: matches\ndata: ${JSON.stringify({ matches: [rowPointer], summary: 'x' })}`,
    );

    expect(event?.type).toBe('matches');
    expect(event?.type === 'matches' && event.data.matches[0]).toMatchObject({
      sheetName: 'Bank Statement',
      row: 444,
      col: 0,
      rowNum: 445,
      endCol: 8,
    });
  });

  it('carries the same pointer fields as the answer event does', () => {
    const fromAnswer = parseSseEventBlock(
      `event: answer\ndata: ${JSON.stringify({ answer: 'x', matches: [rowPointer] })}`,
    );
    const fromMatches = parseSseEventBlock(
      `event: matches\ndata: ${JSON.stringify({ matches: [rowPointer] })}`,
    );

    const a = fromAnswer?.type === 'answer' ? fromAnswer.data.matches?.[0] : undefined;
    const m = fromMatches?.type === 'matches' ? fromMatches.data.matches[0] : undefined;
    expect(m?.endCol).toBe(a?.endCol);
    expect(m?.rowNum).toBe(a?.rowNum);
  });

  it('leaves endCol undefined for a single-cell find pointer', () => {
    const event = parseSseEventBlock(
      `event: matches\ndata: ${JSON.stringify({
        matches: [{ label: '593111', sheetName: 'Bank Statement', row: 4, col: 2, colLetter: 'C', rowNum: 5 }],
      })}`,
    );
    expect(event?.type === 'matches' && event.data.matches[0].endCol).toBeUndefined();
  });
});
