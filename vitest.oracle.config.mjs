import { defineConfig } from 'vitest/config';

// Dedicated config for the Lean-oracle differential tests.
// Kept separate so the default suite never depends on a Lean
// toolchain (the suites also skip when the oracle binary has
// not been built).
export default defineConfig({
  test: {
    environment: 'node',
    include: ['oracle/**/*.test.js'],
    testTimeout: 180_000,
    hookTimeout: 60_000,
  },
});
