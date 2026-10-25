import { baseConfig } from '../eslint.base.js';
import { STDOUT_BAN } from '../eslint.ownership.js';

/*
 * Write's own path-scoped rules (package split, phase 5), which the root held
 * for these files before the move. The reason of the stdout ban lives beside
 * its data in eslint.ownership.js; scripts/check-lint-fixtures.ts proves it
 * still fires here. No ban on auth or read imports: the import check
 * (scripts/check-package-imports.ts) refuses them.
 */

/** @type {import('eslint').Linter.Config[]} */
export default [
  ...baseConfig(import.meta.dirname),
  {
    // stdout is the MCP frame channel: banned in every write file. The bin
    // (src/main.ts) writes only stderr; the CLI builder of core owns stdout.
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
