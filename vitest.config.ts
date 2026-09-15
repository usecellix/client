import { defineConfig, configDefaults } from 'vitest/config';
import path from 'path';

export default defineConfig({
  test: {
    environment: 'node',
    // actionCatalogParity.spec.ts cross-imports cellix_backend's
    // action-catalog.ts, which lives in a separate repo (usecellix/Server)
    // not checked out alongside this one in CI or on Vercel — only on a dev
    // machine with both repos cloned as siblings. Excluded from the default
    // run so `npm test` / `npm run build` stay self-contained for this repo;
    // run it explicitly with `npm run test:parity` from a full monorepo
    // checkout instead.
    exclude: [...configDefaults.exclude, 'src/types/actionCatalogParity.spec.ts'],
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
});
