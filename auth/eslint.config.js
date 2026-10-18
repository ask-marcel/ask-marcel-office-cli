import { baseConfig } from '../eslint.base.js';
import { AUTH_LADDER_IMPORT_BANS, STDOUT_BAN, TOKEN_HELPER_IMPORT_BANS } from '../eslint.ownership.js';

/*
 * Auth's own path-scoped rules (package split, phase 4), which the root held
 * for these files before the move. The reasons of the two import bans live
 * beside their data in eslint.ownership.js, and the reason of the stdout ban
 * beside its rule below; scripts/check-lint-fixtures.ts proves each still
 * fires here.
 */

/** @type {import('eslint').Linter.Config[]} */
export default [
  ...baseConfig(import.meta.dirname),
  {
    // stdout is the MCP frame channel: banned everywhere but the token
    // helper's one JSON line. The bin (src/main.ts) writes only stderr.
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**', 'src/composition/token-helper.ts'],
    rules: {
      'no-restricted-properties': ['error', STDOUT_BAN],
    },
  },
  {
    // The auth ladder never loads the browser.
    files: ['src/infra/auth.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: AUTH_LADDER_IMPORT_BANS }],
    },
  },
  {
    // The token helper never loads the CLI or the MCP server (auth's builders,
    // the MCP SDK, zod), winston, the update notifier, the renderer or,
    // statically, the browser.
    files: ['src/token.ts', 'src/composition/token-helper.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: TOKEN_HELPER_IMPORT_BANS }],
    },
  },
  {
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**'],
  },
];
