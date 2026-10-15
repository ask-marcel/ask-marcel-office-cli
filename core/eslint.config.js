import { baseConfig } from '../eslint.base.js';

/** @type {import('eslint').Linter.Config[]} */
export default [
  ...baseConfig(import.meta.dirname),
  {
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**'],
  },
];
