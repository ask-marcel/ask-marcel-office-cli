import { baseConfig } from '../eslint.base.js';
import { CORE_DYNAMIC_IMPORT_BANS, CORE_FILES, CORE_IMPORT_BANS, CORE_READ_ONLY_FILES, READ_ONLY_IMPORT_BANS, STDOUT_BAN } from '../eslint.ownership.js';

/*
 * Core's own path-scoped rules (package split, phase 3), the root's rules for
 * the files that moved here. The ownership lists name `src/...`, so their globs
 * match core's files from this folder. The reasons live beside each rule in the
 * root eslint.config.js and in eslint.ownership.js; scripts/check-lint-fixtures.ts proves each still fires
 * here.
 */

/** @type {import('eslint').Linter.Config[]} */
export default [
  ...baseConfig(import.meta.dirname),
  {
    // stdout is the MCP frame channel: banned everywhere but the CLI writers.
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**', 'src/composition/cli-shell.ts', 'src/presenter/output.ts'],
    rules: {
      'no-restricted-properties': ['error', STDOUT_BAN],
    },
  },
  {
    // Core never knows which package it serves.
    files: CORE_FILES,
    ignores: ['**/*.test.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: CORE_IMPORT_BANS }],
      'no-restricted-syntax': ['error', ...CORE_DYNAMIC_IMPORT_BANS],
    },
  },
  {
    // The core files that read code also uses never hold the write graph.
    files: CORE_READ_ONLY_FILES,
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: [...READ_ONLY_IMPORT_BANS, ...CORE_IMPORT_BANS] }],
      'no-restricted-syntax': ['error', ...CORE_DYNAMIC_IMPORT_BANS],
    },
  },
  {
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**'],
  },
];
