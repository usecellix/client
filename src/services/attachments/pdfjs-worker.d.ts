/**
 * Vite's `?url` suffix returns the built URL of a file instead of its contents.
 * Declared for the one place it is used rather than pulling in all of
 * `vite/client`, which this project does not otherwise reference.
 */
declare module 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url' {
  const url: string;
  export default url;
}
