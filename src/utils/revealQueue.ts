export function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Visual pacing — slower so each stage is readable; gaps separate step rows from thinking */
export const TIMING = {
  /** Before "Reading…" text appears */
  readingReveal: 950,
  /** After reading label appears, before its loader runs */
  readingSpinner: 1600,
  /** Only the reading step row (no thinking) so the loader + label can be read */
  pauseAfterReadingStepBeforeThinking: 2200,
  /** Minimum time after backend signals "analyzing" before advancing (if still waiting) */
  readingMinRun: 2200,
  /** After analyzing label appears, before its loader runs */
  analyzingSpinner: 1500,
  /** Only the analyzing step row before the thinking panel returns */
  pauseAfterAnalyzingStepBeforeThinking: 2200,
  /** Minimum time analyzing loader runs before composing */
  analyzingMinRun: 1800,
  /** After analyzing step completes, before composing + thought copy */
  pauseBeforeComposing: 900,
  /** After steps complete, before answer card */
  answerReveal: 1200,
  /** After question steps complete */
  questionReveal: 850,
  /** Status row after thinking appears */
  statusReveal: 700,
  /** Thinking hidden: beat before the analyzing step row appears */
  gapBeforeAnalyzingStepRow: 750,
} as const;

export function createGate(): {
  wait: () => Promise<void>;
  open: () => void;
  isOpen: () => boolean;
} {
  let open = false;
  let resolver: (() => void) | null = null;

  return {
    wait: () => {
      if (open) return Promise.resolve();
      return new Promise<void>((resolve) => {
        resolver = resolve;
      });
    },
    open: () => {
      if (open) return;
      open = true;
      resolver?.();
      resolver = null;
    },
    isOpen: () => open,
  };
}

export function waitWithMin(gate: ReturnType<typeof createGate>, minMs: number): Promise<void> {
  return Promise.all([gate.wait(), delay(minMs)]).then(() => undefined);
}
