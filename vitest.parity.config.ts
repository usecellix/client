import { defineConfig } from 'vitest/config';
import path from 'path';

// Separate config for actionCatalogParity.spec.ts, which cross-imports
// cellix_backend's action-catalog.ts from a sibling repo (usecellix/Server).
// The default vitest.config.ts excludes this spec because that import only
// resolves on a dev machine with both repos cloned side by side — never on
// Vercel or in a fresh `usecellix/client`-only clone. Run via `npm run
// test:parity` from a full monorepo checkout.
export default defineConfig({
  test: {
    environment: 'node',
    include: ['src/types/actionCatalogParity.spec.ts'],
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
});
