import { baseConfig } from './eslint.base.js';

/*
 * The root holds no source since the package split (phase 6): it lints its own
 * config files only. Each workspace folder lints itself with its own
 * eslint.config.js (D7), and the path-scoped bans live there, with their data
 * and their reasons in eslint.ownership.js.
 */

/** @type {import('eslint').Linter.Config[]} */
export default [
  ...baseConfig(import.meta.dirname),
  {
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**', 'docs/**', 'scripts/**', '.claude/**', '.agents/**', 'auth/**', 'core/**', 'read/**', 'write/**'],
  },
];
