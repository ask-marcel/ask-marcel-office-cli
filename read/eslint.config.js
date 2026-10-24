import { baseConfig } from '../eslint.base.js';
import { READ_ONLY_IMPORT_BANS, STDOUT_BAN } from '../eslint.ownership.js';

/*
 * Read's own path-scoped rules (package split, phase 6), which the root held
 * for these files before the move. The reasons of the bans live beside their
 * data in eslint.ownership.js; scripts/check-lint-fixtures.ts proves that they
 * fire here. No ban on auth or write imports: the import check
 * (scripts/check-package-imports.ts) refuses them.
 */

/** @type {import('eslint').Linter.Config[]} */
export default [
  ...baseConfig(import.meta.dirname),
  {
    // stdout is the MCP frame channel, and read code never holds the write
    // graph (D9), its type included: both banned in every read file. The CLI
    // builder of core owns stdout.
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**'],
    rules: {
      'no-restricted-properties': ['error', STDOUT_BAN],
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: [READ_ONLY_IMPORT_BANS[0]] }],
    },
  },
  {
    // The skeleton bin writes its answers to stdout until the read bin
    // replaces it.
    files: ['src/main.ts'],
    rules: {
      'no-restricted-properties': 'off',
    },
  },
  {
    // A read file never builds the client of the composition root either:
    // only the composition root (build-deps.ts) takes it from the module that
    // builds it (graph-client.ts).
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**', 'src/composition/build-deps.ts', 'src/infra/graph-client.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: READ_ONLY_IMPORT_BANS }],
    },
  },
  {
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**'],
  },
];
