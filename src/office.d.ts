/**
 * Office.js type definitions.
 *
 * This file previously hand-declared a small subset of the Office/Excel API. That
 * stub silently diverged from the real API — code could reference properties that
 * do not exist at runtime (e.g. Range.row, which is actually Range.rowIndex) and
 * still compile. Point at the shipped definitions instead so the compiler checks
 * our Office.js usage for real.
 *
 * @microsoft/office-js has no "types" entry in its package.json, so it cannot be
 * resolved via tsconfig "types" / `/// <reference types="..." />` and must be
 * referenced by path.
 */
/// <reference path="../node_modules/@microsoft/office-js/dist/office.d.ts" />
