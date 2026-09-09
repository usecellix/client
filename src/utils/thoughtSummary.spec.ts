import { describe, expect, it } from 'vitest';
import { stillWorkingMessage } from './thoughtSummary';

/**
 * Live report (Sept 8, 2026): "user is sitting idle... no loading or
 * anything" during a long, opaque Planner call (30s-4min, nothing for the
 * backend to report until it finishes). `stillWorkingMessage` is the
 * escalating client-side reassurance shown while a wait drags on.
 */
describe('stillWorkingMessage', () => {
  it('shows nothing before the first tier — the existing scripted reveal covers early seconds', () => {
    expect(stillWorkingMessage(0)).toBeNull();
    expect(stillWorkingMessage(5000)).toBeNull();
    expect(stillWorkingMessage(14_999)).toBeNull();
  });

  it('escalates through tiers as elapsed time grows, never regressing to an earlier one', () => {
    const at15s = stillWorkingMessage(15_000);
    const at39s = stillWorkingMessage(39_000);
    const at40s = stillWorkingMessage(40_000);
    const at89s = stillWorkingMessage(89_000);
    const at90s = stillWorkingMessage(90_000);
    const atFiveMin = stillWorkingMessage(5 * 60_000);

    expect(at15s).not.toBeNull();
    expect(at39s).toBe(at15s); // still in the same tier
    expect(at40s).not.toBe(at15s); // escalated
    expect(at89s).toBe(at40s);
    expect(at90s).not.toBe(at40s); // escalated again
    expect(atFiveMin).toBe(at90s); // stays at the highest tier, doesn't invent new ones
  });

  it('never returns an empty or whitespace-only message once a tier is reached', () => {
    for (const elapsed of [15_000, 40_000, 90_000, 600_000]) {
      const message = stillWorkingMessage(elapsed);
      expect(message?.trim().length).toBeGreaterThan(0);
    }
  });
});
