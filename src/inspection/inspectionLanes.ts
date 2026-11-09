/**
 * Client-side deterministic lanes that need Office.js (formulas, names, hidden objects) —
 * things the server never sees. Read-only lanes answer directly; fix lanes return
 * ordinary actions that still go preview → Accept → apply → revert.
 *
 *   P2.4 scan workbook · P3.5 hard-coded numbers (+ move to inputs) · P3.6 broken refs
 *   (+ break links) · P3.13 why #N/A · P4.6 fix #N/A · P3.17 LET rewrite ·
 *   P3.22 explain workbook · P3.24 save sheet as template (headers only)
 */
import type { SheetAction } from '@/types/sheet-actions';
import { brokenRefsReport, explainWorkbook, hardcodedReport, scanReport, scanWorkbook, type WorkbookScan } from './workbookScan';
import { diagnoseLookupNA, INPUTS_SHEET, planBreakLinks, planMoveLiteralsToInputs, toLet } from './formulaTools';
import { auditEntries, downloadAuditCsv, downloadAuditPdf } from '@/services/localAuditLog';
import { listUserTemplates, saveUserTemplate } from './userTemplates';

/* global Excel */

export interface InspectionResult {
  answer: string;
  actions?: SheetAction[];
}

export type InspectionKind =
  | 'scan'
  | 'hardcoded'
  | 'hardcoded-fix'
  | 'broken-refs'
  | 'break-links'
  | 'lookup-why'
  | 'lookup-fix'
  | 'let'
  | 'explain-workbook'
  | 'save-template'
  | 'use-template'
  | 'audit-export';

export function detectInspection(message: string): InspectionKind | null {
  const m = message.trim();
  if (!m) return null;
  if (/\b(move|fix|replace|convert)\b.*\bhard[-\s]?coded\b|\bhard[-\s]?coded\b.*\b(to|into)\s+(named\s+)?inputs?\b/i.test(m)) return 'hardcoded-fix';
  if (/\bhard[-\s]?coded\b|\bhardcoded\b|\bmagic numbers?\b|\bconstants?\s+in\s+(my\s+)?formulas\b/i.test(m)) return 'hardcoded';
  if (/\bbreak\b.*\b(links?|references?|refs?)\b/i.test(m)) return 'break-links';
  if (/\b(broken|external)\s+(references?|refs?|links?)\b|#REF!/i.test(m)) return 'broken-refs';
  if (/\bfix\b.*#N\/A|#N\/A.*\bfix\b/i.test(m)) return 'lookup-fix';
  if (/#N\/A/i.test(m) && /\b(why|what|reason|diagnos\w*|wrong|showing)\b/i.test(m)) return 'lookup-why';
  if (/\bLET\b/.test(m) && /\b(rewrite|refactor|convert|use|using)\b/i.test(m)) return 'let';
  if (/\b(scan|audit|check)\b.*\b(workbook|file)\b.*\b(problems?|issues?|errors?)?\b/i.test(m) && /\b(scan|problems?|issues?|health)\b/i.test(m)) return 'scan';
  if (/\b(explain|summari[sz]e|describe|walk me through)\b.*\b(workbook|file)\b|\bwhat('s| is)\s+in\s+(this|the|my)\s+workbook\b/i.test(m)) return 'explain-workbook';
  if (/\b(export|download)\b.*\baudit\b|\baudit\s+(trail|log)\b.*\b(export|download)\b/i.test(m)) return 'audit-export';
  if (/\bsave\b.*\bas\s+(a\s+)?(my\s+)?template\b/i.test(m)) return 'save-template';
  if (/\b(create|use|new)\b.*\b(from\s+)?(my\s+)?template\b/i.test(m) && listUserTemplates().some((t) => m.toLowerCase().includes(t.name.toLowerCase()))) return 'use-template';
  return null;
}

/** Whether this kind proposes a write (needs Action mode). */
export const inspectionWrites = (kind: InspectionKind): boolean =>
  ['hardcoded-fix', 'break-links', 'lookup-fix', 'let', 'use-template'].includes(kind);

const columnFromMessage = (m: string): number | undefined => {
  const hit = /\bcolumn\s+([A-Za-z])\b/i.exec(m);
  return hit ? hit[1].toUpperCase().charCodeAt(0) - 65 : undefined;
};

async function readActiveSheet(): Promise<{ name: string; formulas: unknown[][]; values: unknown[][] }> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load('name');
    const used = ws.getUsedRangeOrNullObject(true);
    used.load(['isNullObject', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
    await ctx.sync();
    if (used.isNullObject) return { name: ws.name, formulas: [], values: [] };
    const block = ws.getRangeByIndexes(0, 0, used.rowIndex + used.rowCount, used.columnIndex + used.columnCount);
    block.load(['formulas', 'values']);
    await ctx.sync();
    return { name: ws.name, formulas: block.formulas as unknown[][], values: block.values as unknown[][] };
  });
}

export async function runInspection(kind: InspectionKind, message: string): Promise<InspectionResult> {
  if (kind === 'audit-export') {
    const list = auditEntries();
    if (!list.length) return { answer: 'No accepted actions in this browser yet, so the audit trail is empty.' };
    const asPdf = /\bpdf\b/i.test(message);
    if (asPdf) downloadAuditPdf();
    else downloadAuditCsv();
    const applied = list.filter((e) => e.status === 'Applied').length;
    return {
      answer: `Downloaded **cellix-audit-trail.${asPdf ? 'pdf' : 'csv'}**${asPdf ? '' : ' (ask *…as PDF* for a printable version)'} — ${list.length} entr${list.length === 1 ? 'y' : 'ies'} (${applied} applied, ${list.length - applied} reverted): time, change set, what was asked, action, sheet, range, cell count. Built in this browser; it contains **no cell values** and nothing was fetched from the server.`,
    };
  }
  if (kind === 'save-template') {
    const sheet = await readActiveSheet();
    const headers = (sheet.values[0] ?? []).map((v) => String(v ?? '')).filter(Boolean);
    const name = /\btemplate\s+(?:named|called)\s+["“']?([^"”']+)["”']?/i.exec(message)?.[1]?.trim() || sheet.name;
    const ok = saveUserTemplate({ name, headers });
    return {
      answer: ok
        ? `Saved **${name}** as your template: ${headers.join(' · ')}. Only the header row is saved (in this browser) — no data. Use it with *Create from template ${name}*.`
        : 'Could not save the template in this browser (storage is blocked).',
    };
  }
  if (kind === 'use-template') {
    const t = listUserTemplates().find((x) => message.toLowerCase().includes(x.name.toLowerCase()))!;
    return {
      answer: `I'll create a sheet **${t.name}** with your saved headers: ${t.headers.join(' · ')}.`,
      actions: [
        { type: 'ADD_SHEET', name: t.name } as SheetAction,
        { type: 'WRITE_TABLE', sheetName: t.name, headers: t.headers, rows: [] } as SheetAction,
      ],
    };
  }

  if (kind === 'lookup-why' || kind === 'lookup-fix') {
    const sheet = await readActiveSheet();
    const d = diagnoseLookupNA(sheet.formulas, sheet.values, columnFromMessage(message));
    if (!d) return { answer: `No #N/A results found on **${sheet.name}**.` };
    if (kind === 'lookup-why' || !d.fix?.length) {
      return {
        answer: `${d.message}\n\nAffected: ${d.naCells.join(', ')}.${d.fix?.length ? `\n\nFix: ask *Fix the #N/A in column ${d.column}* — I'll convert ${d.fix.length} lookup value(s) in column ${d.lookupRange} to numbers (preview first).` : ''}`,
      };
    }
    return {
      answer: `${d.message}\n\n**Fix:** convert ${d.fix.length} cell(s) in column ${d.lookupRange} from text to numbers: ${d.fix.map((f) => `${f.address} "${f.value}" → ${f.value}`).join(', ')}. The lookup formulas are not changed; #N/A should disappear after Accept.`,
      actions: [{ type: 'BATCH_SET', sheetName: sheet.name, operations: d.fix.map((f) => ({ address: f.address, value: f.value })) } as SheetAction],
    };
  }

  if (kind === 'let') {
    const sheet = await readActiveSheet();
    const target = /\b([A-Za-z]{1,2})(\d+)\b/.exec(message.replace(/\bLET\b/g, ''));
    const col = target ? target[1].toUpperCase().charCodeAt(0) - 65 : -1;
    const row = target ? Number(target[2]) - 1 : -1;
    const first = String(sheet.formulas[row]?.[col] ?? '');
    if (!first.startsWith('=')) return { answer: `${target ? target[0].toUpperCase() : 'That cell'} has no formula to rewrite.` };
    const labels: Record<string, string> = {};
    (sheet.values[0] ?? []).forEach((h, c) => {
      if (h) labels[`${String.fromCharCode(65 + c)}${row + 1}`] = String(h).replace(/\s+(.)/g, (_, x: string) => x.toUpperCase());
    });
    const ops: Array<{ address: string; formula: string }> = [];
    let allEquivalent = true;
    let sample = toLet(first, labels);
    if (!sample) return { answer: `\`${first}\` has nothing worth naming (no repeated reference or constant), so LET would not make it clearer.` };
    for (let r = row; r < sheet.formulas.length; r += 1) {
      const f = String(sheet.formulas[r]?.[col] ?? '');
      if (!f.startsWith('=')) break;
      const rowLabels: Record<string, string> = {};
      (sheet.values[0] ?? []).forEach((h, c) => {
        if (h) rowLabels[`${String.fromCharCode(65 + c)}${r + 1}`] = String(h).replace(/\s+(.)/g, (_, x: string) => x.toUpperCase());
      });
      const rw = toLet(f, rowLabels);
      if (!rw) break;
      allEquivalent = allEquivalent && rw.equivalent;
      ops.push({ address: `${String.fromCharCode(65 + col)}${r + 1}`, formula: rw.rewritten });
      if (r === row) sample = rw;
    }
    if (!allEquivalent) return { answer: 'The LET version did not expand back to the original formula on every row, so I will not propose it.' };
    return {
      answer: [
        `**LET rewrite** for ${ops.length} row(s) of column ${String.fromCharCode(65 + col)}:`,
        `- Before: \`${sample.original}\``,
        `- After: \`${sample.rewritten}\``,
        `- Names: ${sample.names.map((n) => `${n.name} = ${n.expr}`).join(', ')}`,
        `**Equivalence check:** expanding the names back gives the original formula exactly on all ${ops.length} rows, so the results are identical. Needs Excel 2021/365 (LET).`,
      ].join('\n'),
      actions: [{ type: 'BATCH_SET', sheetName: sheet.name, operations: ops } as SheetAction],
    };
  }

  const scan: WorkbookScan = await scanWorkbook();
  if (kind === 'scan') return { answer: scanReport(scan) };
  if (kind === 'explain-workbook') return { answer: explainWorkbook(scan) };
  if (kind === 'hardcoded') return { answer: hardcodedReport(scan) };
  if (kind === 'broken-refs') return { answer: brokenRefsReport(scan) };
  if (kind === 'hardcoded-fix') {
    const findings = scan.sheets.filter((s) => s.name !== INPUTS_SHEET).flatMap((s) => s.hardcoded);
    if (!findings.length) return { answer: 'No hard-coded numbers to move.' };
    const inputs = scan.sheets.find((s) => s.name === INPUTS_SHEET);
    const plan = planMoveLiteralsToInputs(findings, inputs?.rows ?? 0);
    return {
      answer: [`Moving **${plan.lines.length}** constant(s) to the **${INPUTS_SHEET}** sheet; each formula then points at its input cell:`, ...plan.lines].join('\n'),
      actions: plan.actions,
    };
  }
  // break-links
  const plan = planBreakLinks(scan);
  if (!plan.actions.length) return { answer: 'Nothing to break — no #REF! or external-link formulas found.' };
  const badNames = scan.names.filter((n) => n.broken);
  return {
    answer: [
      `**Breaking ${plan.lines.length} link(s)** — each formula is replaced by what it shows now:`,
      ...plan.lines,
      ...(badNames.length ? ['', `Broken names (${badNames.map((n) => n.name).join(', ')}) must be deleted in Formulas → Name Manager — Office.js cannot delete them safely here.`] : []),
    ].join('\n'),
    actions: plan.actions,
  };
}
