import { baseConfig } from '../eslint.base.js';
import { STDOUT_BAN } from '../eslint.ownership.js';

/*
 * The facade's own path-scoped rule (ADR 0005). Each bin of this folder starts
 * the bin of one package in the same process, and that bin can be an MCP
 * server. The reason of the stdout ban lives beside its data in
 * eslint.ownership.js; scripts/check-lint-fixtures.ts proves it still fires
 * here.
 */

/** @type {import('eslint').Linter.Config[]} */
export default [
  ...baseConfig(import.meta.dirname),
  {
    // stdout is the MCP frame channel of the bin that a facade bin starts:
    // banned in every facade file. A facade bin writes only stderr.
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**'],
    rules: {
      'no-restricted-properties': ['error', STDOUT_BAN],
    },
  },
  {
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**'],
  },
];
