/**
 * P3.10 / A16.8 — audit trail of EVERY accepted action, kept in the add-in and exported
 * as a file generated locally. It records what was done (time, sheet, range, action type,
 * cell count, undo) — never cell values, and nothing is fetched from the server to build it.
 *
 * Kept in memory for the session and mirrored to this browser's storage (per-viewer
 * convenience; guarded — storage may be blocked).
 */
import type { SheetAction } from '@/types/sheet-actions';
import { buildTablePdf, downloadPdf } from './simplePdf';

export interface AuditEntry {
  at: string;
  changeSetId: string;
  summary: string;
  actionType: string;
  sheet: string;
  range: string;
  cells: number;
  status: 'Applied' | 'Reverted';
}

const KEY = 'cellix.localAudit.v1';
const MAX = 2000;
let entries: AuditEntry[] = load();

function load(): AuditEntry[] {
  try {
    const raw = globalThis.localStorage?.getItem(KEY);
    const parsed = raw ? (JSON.parse(raw) as unknown) : [];
    return Array.isArray(parsed) ? (parsed as AuditEntry[]) : [];
  } catch {
    return [];
  }
}

function persist(): void {
  try {
    globalThis.localStorage?.setItem(KEY, JSON.stringify(entries.slice(-MAX)));
  } catch {
    /* storage blocked: the in-memory log still works for this session */
  }
}

const cellsInRange = (range: string): number => {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(range.replace(/^.*!/, ''));
  if (!m) return 0;
  const col = (l: string) => [...l.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  if (!m[3]) return 1;
  return (Math.abs(col(m[3]) - col(m[1])) + 1) * (Math.abs(Number(m[4]) - Number(m[2])) + 1);
};

/** Describe one action WITHOUT its values: where it acted and how many cells. */
export function describeAction(action: SheetAction): { sheet: string; range: string; cells: number } {
  const a = action as unknown as Record<string, unknown>;
  const sheet = String(a.sheetName ?? a.name ?? a.destSheet ?? a.sourceSheet ?? '');
  if (Array.isArray(a.operations)) {
    const addrs = (a.operations as Array<{ address?: string }>).map((o) => String(o.address ?? ''));
    return { sheet, range: addrs.length > 3 ? `${addrs[0]}…${addrs[addrs.length - 1]}` : addrs.join(', '), cells: addrs.length };
  }
  if (Array.isArray(a.headers)) {
    const rows = Array.isArray(a.rows) ? (a.rows as unknown[]).length : 0;
    return { sheet, range: 'from A1', cells: (rows + 1) * (a.headers as unknown[]).length };
  }
  const range = String(a.range ?? a.targetRange ?? a.address ?? '');
  return { sheet, range, cells: range ? cellsInRange(range) : 0 };
}

export function recordApplied(changeSetId: string | undefined, summary: string, actions: SheetAction[], at = new Date()): void {
  for (const action of actions) {
    const d = describeAction(action);
    entries.push({
      at: at.toISOString(),
      changeSetId: changeSetId ?? '',
      summary: summary.split('\n')[0].slice(0, 200),
      actionType: action.type,
      ...d,
      status: 'Applied',
    });
  }
  persist();
}

export function recordReverted(changeSetId: string): void {
  let hit = false;
  entries = entries.map((e) => (e.changeSetId === changeSetId ? ((hit = true), { ...e, status: 'Reverted' as const }) : e));
  if (hit) persist();
}

export const auditEntries = (): AuditEntry[] => [...entries];

export function clearAuditLog(): void {
  entries = [];
  persist();
}

const csvCell = (v: unknown) => {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};

export function auditCsv(list: AuditEntry[] = entries): string {
  const head = ['Time (UTC)', 'Change set', 'What was asked / done', 'Action', 'Sheet', 'Range', 'Cells', 'Status'];
  const rows = list.map((e) => [e.at, e.changeSetId, e.summary, e.actionType, e.sheet, e.range, e.cells, e.status]);
  return [head, ...rows].map((r) => r.map(csvCell).join(',')).join('\r\n');
}

/** Builds the file in the browser and hands it to the user — no server round trip. */
export function downloadAuditCsv(): number {
  const csv = auditCsv();
  const blob = new Blob([`﻿${csv}`], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `cellix-audit-trail-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  return entries.length;
}

/** The audit trail as a PDF (A4 landscape) — built in the browser, no cell values. */
export function auditPdf(list: AuditEntry[] = entries, generatedAt = new Date()): string {
  const applied = list.filter((e) => e.status === 'Applied').length;
  return buildTablePdf(
    {
      title: 'CELLIX - Audit trail of accepted actions',
      subtitle: `${list.length} entries (${applied} applied, ${list.length - applied} reverted). Records what was done, where and when - no cell values.`,
      columns: [
        { header: 'Time (UTC)', width: 92 },
        { header: 'Change set', width: 70 },
        { header: 'What was asked / done', width: 250 },
        { header: 'Action', width: 90 },
        { header: 'Sheet', width: 90 },
        { header: 'Range', width: 90 },
        { header: 'Cells', width: 34 },
        { header: 'Status', width: 54 },
      ],
      rows: list.map((e) => [e.at.replace('T', ' ').slice(0, 19), e.changeSetId.slice(0, 8), e.summary, e.actionType, e.sheet, e.range, String(e.cells), e.status]),
      footer: 'CELLIX audit trail - working paper record. Reviewed by: ____________',
    },
    generatedAt,
  );
}

export function downloadAuditPdf(): number {
  downloadPdf(auditPdf(), `cellix-audit-trail-${new Date().toISOString().slice(0, 10)}.pdf`);
  return entries.length;
}
