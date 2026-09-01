import { describe, expect, it, vi, afterEach, beforeEach } from 'vitest';
import {
  CAPABILITY_PROBE_FORMULA,
  probeExcelCapabilities,
  resetCapabilityCache,
} from './capabilityProbe';

/* global globalThis */

/**
 * TASKS.md #152 — probe before trusting a formula family.
 *
 * Shortcut wrote `=VSTACK("a","b")` to a scratch cell before relying on dynamic
 * arrays. Cellix shipped #142's consolidation formula as an assumption: on
 * pre-365 Excel every function in it is `#NAME?` and nothing noticed.
 */
function installExcel(probeResult: unknown, opts: { throwOnRun?: boolean } = {}) {
  const cleared: string[] = [];
  const written: string[] = [];

  (globalThis as Record<string, unknown>).Excel = {
    ClearApplyTo: { all: 'All' },
    run: async (cb: (ctx: unknown) => Promise<unknown>) => {
      if (opts.throwOnRun) throw new Error('Excel is in cell-editing mode.');
      const cell = {
        set formulas(v: unknown) { written.push(String((v as string[][])[0][0])); },
        load: vi.fn(),
        values: [[probeResult]],
        text: [[String(probeResult)]],
        clear: (what: string) => cleared.push(what),
      };
      return cb({
        workbook: { worksheets: { getActiveWorksheet: () => ({ getRange: () => cell }) } },
        sync: vi.fn(async () => undefined),
      });
    },
  };
  return { cleared, written };
}

describe('probeExcelCapabilities', () => {
  const original = (globalThis as Record<string, unknown>).Excel;
  beforeEach(() => resetCapabilityCache());
  afterEach(() => {
    (globalThis as Record<string, unknown>).Excel = original;
    resetCapabilityCache();
    vi.restoreAllMocks();
  });

  it('reports support when the probe evaluates to a number', async () => {
    installExcel(2);
    const caps = await probeExcelCapabilities();
    expect(caps).toEqual({ dynamicArrays: true, probed: true });
  });

  it('reports NO support when the probe returns #NAME?', async () => {
    installExcel('#NAME?');
    const caps = await probeExcelCapabilities();
    expect(caps.dynamicArrays).toBe(false);
    expect(caps.probed).toBe(true);
    expect(caps.reason).toContain('#NAME?');
  });

  it('exercises the exact function set the consolidation formula uses', () => {
    // A pass must mean "that specific formula will work", not "some dynamic
    // arrays exist" — so the probe names every function #142 depends on.
    for (const fn of ['LET', 'VSTACK', 'HSTACK', 'FILTER', 'BYROW', 'LAMBDA', 'EXPAND', 'DROP']) {
      expect(CAPABILITY_PROBE_FORMULA).toContain(fn);
    }
  });

  it('is non-spilling — wrapped in COUNT so it can never overwrite a neighbour', () => {
    expect(CAPABILITY_PROBE_FORMULA.startsWith('=COUNT(')).toBe(true);
  });

  it('always clears the scratch cell, leaving no trace', async () => {
    const { cleared, written } = installExcel(2);
    await probeExcelCapabilities();
    expect(written).toEqual([CAPABILITY_PROBE_FORMULA]);
    expect(cleared).toEqual(['All']);
  });

  it('caches per session — Excel cannot gain functions mid-session', async () => {
    const first = installExcel(2);
    await probeExcelCapabilities();
    await probeExcelCapabilities();
    await probeExcelCapabilities();
    expect(first.written).toHaveLength(1);
  });

  it('treats a failed probe as UNKNOWN, never as supported', async () => {
    installExcel(null, { throwOnRun: true });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    const caps = await probeExcelCapabilities();
    // probed:false is the signal the server keys off to take the safe path.
    expect(caps.probed).toBe(false);
    expect(caps.dynamicArrays).toBe(false);
    warn.mockRestore();
  });

  it('does not blow up outside a task pane', async () => {
    delete (globalThis as Record<string, unknown>).Excel;
    const caps = await probeExcelCapabilities();
    expect(caps.probed).toBe(false);
    expect(caps.reason).toContain('Office.js unavailable');
  });
});
