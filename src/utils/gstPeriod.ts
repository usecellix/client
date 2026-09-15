/**
 * Resolve a GST reconciliation period from free-text chat phrasing into a concrete
 * inclusive date range (yyyy-mm-dd), so books/portal rows can be filtered to that
 * period before the matching pipeline runs.
 *
 * Deliberately conservative: returns null (no filtering) whenever the phrasing is
 * ambiguous rather than guessing — a wrong silent filter is worse than no filter.
 */

export interface ResolvedGstPeriod {
  /** Human-readable label for chat/summary text, e.g. "April 2024" or "Q1 FY25". */
  label: string;
  /** Inclusive start date, yyyy-mm-dd. */
  start: string;
  /** Inclusive end date, yyyy-mm-dd. */
  end: string;
}

const MONTH_LABELS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const MONTH_RE_SRC =
  '(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)';

function monthIndexFromName(name: string): number {
  const short = name.slice(0, 3).toLowerCase();
  return ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].indexOf(
    short,
  );
}

function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

/** 2-digit years are always post-GST (2017+), so 2-digit "24" always means 2024. */
function expandYear(raw: string): number {
  return raw.length >= 4 ? Number(raw) : 2000 + Number(raw);
}

function monthRange(year: number, monthIdx0: number): { start: string; end: string } {
  const start = `${year}-${pad2(monthIdx0 + 1)}-01`;
  const lastDay = new Date(Date.UTC(year, monthIdx0 + 1, 0)).getUTCDate();
  const end = `${year}-${pad2(monthIdx0 + 1)}-${pad2(lastDay)}`;
  return { start, end };
}

type Quarter = 1 | 2 | 3 | 4;

/** Indian FY quarters: Q1 Apr-Jun, Q2 Jul-Sep, Q3 Oct-Dec, Q4 Jan-Mar (falls in fyStartYear+1). */
function quarterRange(fyStartYear: number, quarter: Quarter): { start: string; end: string; label: string } {
  const startMonth0ByQ: Record<Quarter, number> = { 1: 3, 2: 6, 3: 9, 4: 0 };
  const startYear = quarter === 4 ? fyStartYear + 1 : fyStartYear;
  const startMonth0 = startMonth0ByQ[quarter];
  const start = monthRange(startYear, startMonth0).start;
  const endMonthAbs = startMonth0 + 2;
  const endYear = startYear + Math.floor(endMonthAbs / 12);
  const endMonth0 = endMonthAbs % 12;
  const end = monthRange(endYear, endMonth0).end;
  const label = `Q${quarter} FY${String((fyStartYear + 1) % 100).padStart(2, '0')}`;
  return { start, end, label };
}

function fyQuarterOf(d: Date): { fyStartYear: number; quarter: Quarter } {
  const m = d.getMonth();
  if (m >= 3) {
    const quarter = (Math.floor((m - 3) / 3) + 1) as Quarter;
    return { fyStartYear: d.getFullYear(), quarter };
  }
  return { fyStartYear: d.getFullYear() - 1, quarter: 4 };
}

/**
 * Extract a confident period from free text. Returns null when nothing recognizable
 * matches, or when the reference is genuinely ambiguous — callers should fall back to
 * no date filtering in that case, never guess.
 */
export function resolveGstPeriod(text: string, now: Date = new Date()): ResolvedGstPeriod | null {
  const t = text.toLowerCase();

  if (/\blast\s+quarter\b/.test(t)) {
    const cur = fyQuarterOf(now);
    const isQ1 = cur.quarter === 1;
    const prevQuarter = (isQ1 ? 4 : cur.quarter - 1) as Quarter;
    const prevFyStart = isQ1 ? cur.fyStartYear - 1 : cur.fyStartYear;
    const r = quarterRange(prevFyStart, prevQuarter);
    return { label: `last quarter (${r.label})`, start: r.start, end: r.end };
  }
  if (/\bthis\s+quarter\b/.test(t)) {
    const cur = fyQuarterOf(now);
    const r = quarterRange(cur.fyStartYear, cur.quarter);
    return { label: `this quarter (${r.label})`, start: r.start, end: r.end };
  }
  if (/\blast\s+month\b/.test(t)) {
    const rawMonth = now.getMonth() - 1;
    const year = rawMonth < 0 ? now.getFullYear() - 1 : now.getFullYear();
    const monthIdx0 = (rawMonth + 12) % 12;
    const r = monthRange(year, monthIdx0);
    return { label: `${MONTH_LABELS[monthIdx0]} ${year}`, ...r };
  }
  if (/\bthis\s+month\b/.test(t)) {
    const r = monthRange(now.getFullYear(), now.getMonth());
    return { label: `${MONTH_LABELS[now.getMonth()]} ${now.getFullYear()}`, ...r };
  }

  // "Q1 FY25" / "Q1 FY 2024-25" / "Q1 FY24-25" — requires an explicit FY marker to stay unambiguous.
  const qMatch = t.match(/\bq([1-4])\s*(?:fy|f\.?y\.?)\s*(\d{2,4})(?:\s*[-/]\s*(\d{2,4}))?\b/i);
  if (qMatch) {
    const quarter = Number(qMatch[1]) as Quarter;
    const fyStartYear = qMatch[3] ? expandYear(qMatch[2]) : expandYear(qMatch[2]) - 1;
    const r = quarterRange(fyStartYear, quarter);
    return { label: r.label, start: r.start, end: r.end };
  }

  // Month name + year: "April 2024", "Apr 2024", "Apr-24", "Apr24", "Apr 24".
  const monthMatch = t.match(new RegExp(`\\b${MONTH_RE_SRC}[\\s.-]{0,2}(\\d{2,4})\\b`, 'i'));
  if (monthMatch) {
    const monthIdx0 = monthIndexFromName(monthMatch[1]);
    if (monthIdx0 >= 0) {
      const year = expandYear(monthMatch[2]);
      const r = monthRange(year, monthIdx0);
      return { label: `${MONTH_LABELS[monthIdx0]} ${year}`, ...r };
    }
  }

  // "2024-04" or "2024/04" (yyyy-mm).
  const ymMatch = t.match(/\b(20\d{2})[-/](0?[1-9]|1[0-2])\b/);
  if (ymMatch) {
    const year = Number(ymMatch[1]);
    const monthIdx0 = Number(ymMatch[2]) - 1;
    const r = monthRange(year, monthIdx0);
    return { label: `${MONTH_LABELS[monthIdx0]} ${year}`, ...r };
  }

  // "04/2024" or "04-2024" (mm/yyyy) — must not match inside a full dd/mm/yyyy date.
  const myMatch = t.match(/(?<!\d[/-])\b(0?[1-9]|1[0-2])[/-](\d{4})\b(?!\d)/);
  if (myMatch) {
    const monthIdx0 = Number(myMatch[1]) - 1;
    const year = Number(myMatch[2]);
    const r = monthRange(year, monthIdx0);
    return { label: `${MONTH_LABELS[monthIdx0]} ${year}`, ...r };
  }

  return null;
}
