import { describe, expect, it } from 'vitest';
import { formatFullDateTime, formatRelativeTime } from './conversationTurn';

describe('formatRelativeTime', () => {
  const now = new Date('2026-08-28T12:00:00Z');

  it('shows "just now" for anything under 45 seconds old', () => {
    expect(formatRelativeTime(new Date('2026-08-28T11:59:30Z'), now)).toBe('just now');
  });

  it('shows minutes for under an hour', () => {
    expect(formatRelativeTime(new Date('2026-08-28T11:55:00Z'), now)).toBe('5m ago');
  });

  it('shows hours for under a day', () => {
    expect(formatRelativeTime(new Date('2026-08-28T09:00:00Z'), now)).toBe('3h ago');
  });

  it('shows days for under a month', () => {
    expect(formatRelativeTime(new Date('2026-08-23T12:00:00Z'), now)).toBe('5d ago');
  });

  it('shows months for under a year', () => {
    expect(formatRelativeTime(new Date('2026-06-28T12:00:00Z'), now)).toBe('2mo ago');
  });

  it('shows years beyond that', () => {
    expect(formatRelativeTime(new Date('2024-08-28T12:00:00Z'), now)).toBe('2y ago');
  });

  it('never goes negative for a timestamp slightly in the future (clock skew)', () => {
    expect(formatRelativeTime(new Date('2026-08-28T12:00:05Z'), now)).toBe('just now');
  });
});

describe('formatFullDateTime', () => {
  it('renders a human-readable full date and time for the hover tooltip', () => {
    const result = formatFullDateTime(new Date('2026-08-23T12:18:00Z'));
    expect(result).toContain('2026');
    expect(result).toMatch(/Aug/);
  });
});
