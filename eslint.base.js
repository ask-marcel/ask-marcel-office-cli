import pluginJs from '@eslint/js';
import prettier from 'eslint-plugin-prettier';
import securityPlugin from 'eslint-plugin-security';
import sonarjsPlugin from 'eslint-plugin-sonarjs';
import unicornPlugin from 'eslint-plugin-unicorn';
import globals from 'globals';
import tsPlugin from 'typescript-eslint';

/*
 * The lint rules every folder shares (package split, D7): the root config and
 * the config of each workspace folder (core/, auth/, read/, write/) spread this
 * base, then add the blocks that only their own files need. Globs here are
 * relative to the config that spreads the base, so `src/**` names core/src/
 * in core's config and read/src/ in read's. The root has no src/ (phase 6).
 *
 * tsconfigRootDir: the folder whose tsconfig.json the type-aware lint
 * (LINT_STRICT) reads; pass `import.meta.dirname` of the spreading config.
 */

// Rule 13: the whole mocking surface of bun:test. `mock` is process-global once
// installed and leaks into every other test file; `spyOn`, `jest` and `vi` are
// the same machinery under other names.
const MOCK_BAN = {
  name: 'bun:test',
  importNames: ['mock', 'spyOn', 'jest', 'vi'],
  message:
    '`mock`, `spyOn`, `jest` and `vi` from bun:test are forbidden: mocks leak across test files. Use dependency injection: refactor the production code to accept the SDK as a parameter, then pass a fake at construction (hard rule 13).',
};

// Rule 13, the other half: a call-recording assertion, whatever built the spy.
// ESLint replaces the options of `no-restricted-syntax` when a second block
// matches the same file, so a config with its own `no-restricted-syntax` block
// (core) spreads this list into it.
export const CALL_ASSERTION_BANS = [
  {
    selector: 'MemberExpression[property.name=/^(toHaveBeen(Last|Nth)?Called|toBeCalled|toHave(Last|Nth)?Returned)/]',
    message: 'No call-recording assertions: a hand-written fake records what it received and the test asserts on that record (hard rule 13).',
  },
];

// The shared test fakes (core/src/test-helpers, reached through core's dev-only
// `./test-helpers/*` export, and each folder's own src/test-helpers) are for
// tests only: no build may carry one (package split, Target layout).
const TEST_HELPERS_BAN = {
  regex: '(^|/)test-helpers/',
  message: 'A test fake is dev-only: production code never imports test-helpers/. Take the real port, and pass the fake from the test.',
};

/** @type {(tsconfigRootDir: string) => import('eslint').Linter.Config[]} */
export const baseConfig = (tsconfigRootDir) => [
  pluginJs.configs.recommended,
  ...tsPlugin.configs.recommended,
  securityPlugin.configs.recommended,
  {
    // Rule 15: no inline ignore, ever. A directive comment (eslint-disable*,
    // eslint-enable, global, exported) has no effect, and ESLint reports each
    // one, so --max-warnings=0 fails on the comment, and the finding that it
    // tried to hide shows beside it.
    linterOptions: { noInlineConfig: true },
  },
  {
    languageOptions: { globals: globals.node },
  },
  {
    files: ['**/*.ts'],
    rules: {
      'func-style': ['error', 'expression'],
      'no-console': ['error'],
      'prefer-template': 'error',
      quotes: ['error', 'single', { avoidEscape: true }],
      'no-restricted-imports': ['error', { paths: [MOCK_BAN] }],
      'no-restricted-syntax': ['error', ...CALL_ASSERTION_BANS],
      // Rule 15, the ignores of other tools: every @ts- form (the recommended
      // preset allows a @ts-expect-error with a description), and the markers
      // that other tools read, anywhere in a comment.
      '@typescript-eslint/ban-ts-comment': ['error', { 'ts-expect-error': true, 'ts-ignore': true, 'ts-nocheck': true, 'ts-check': false }],
      'no-warning-comments': [
        'error',
        {
          terms: [
            'prettier-ignore',
            'stryker disable',
            'nosonar',
            'sonar-ignore',
            'snyk-ignore',
            'deepcode ignore',
            'biome-ignore',
            'oxlint-disable',
            'c8 ignore',
            'v8 ignore',
            'istanbul ignore',
            'gitleaks:allow',
          ],
          location: 'anywhere',
        },
      ],
      '@typescript-eslint/explicit-function-return-type': ['error', { allowExpressions: true, allowTypedFunctionExpressions: true }],
      '@typescript-eslint/consistent-type-definitions': ['error', 'type'],
      // The rule ships no default argsIgnorePattern, so it flags deliberately-unused
      // contract parameters (e.g. the local-only commands' redirect `execute(_graph,
      // _params)` shims, where the Command type forces both params but the local
      // implementation uses neither). Allow a leading underscore to mark them.
      // Atelier rule 15: project-level severity/config, never an inline ignore.
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
    },
  },
  {
    // The same rule with the test-helpers ban added: ESLint replaces a rule's
    // options when two blocks match a file, so this block repeats the mock ban.
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**'],
    rules: {
      'no-restricted-imports': ['error', { paths: [MOCK_BAN], patterns: [TEST_HELPERS_BAN] }],
    },
  },
  ...(process.env['LINT_STRICT']
    ? [
        {
          files: ['src/**/*.ts'],
          languageOptions: {
            parserOptions: {
              projectService: true,
              tsconfigRootDir,
            },
          },
          rules: {
            '@typescript-eslint/no-unnecessary-type-assertion': 'error',
            '@typescript-eslint/prefer-promise-reject-errors': 'error',
          },
        },
      ]
    : []),
  {
    plugins: { prettier },
    rules: {
      'prettier/prettier': [
        1,
        {
          endOfLine: 'lf',
          printWidth: 180,
          semi: true,
          singleQuote: true,
          tabWidth: 2,
          trailingComma: 'es5',
        },
      ],
    },
  },
  {
    plugins: { unicorn: unicornPlugin },
    rules: {
      'unicorn/empty-brace-spaces': 'off',
      'unicorn/no-null': 'off',
    },
  },
  {
    rules: {
      'security/detect-object-injection': 'off',
      'security/detect-unsafe-regex': 'off',
      'security/detect-non-literal-fs-filename': 'off',
    },
  },
  sonarjsPlugin.configs.recommended,
  {
    rules: {
      'sonarjs/no-unused-vars': 'off',
      'sonarjs/no-empty-test-file': 'off',
      'sonarjs/cognitive-complexity': 'off',
      // Disabled at the project level (atelier rule 15): TypeScript's strict
      // mode + SonarJS recommended produces false positives on idiomatic
      // patterns this codebase uses heavily.
      // - `no-useless-intersection`: misfires on branded-type intersections
      //   (`string & { __brand }`), the canonical atelier pattern.
      // - `function-return-type`: misfires on factory functions that return
      //   `Result<T, E>` since TS sees the union as multiple return shapes.
      // - `null-dereference`: TypeScript itself enforces strict null checks;
      //   SonarJS duplicates and routinely misfires (e.g., on `Object.keys()`
      //   loop variables, `String.prototype.split` results, regex matches).
      // - `different-types-comparison`: misfires whenever runtime-defensive
      //   `=== undefined` / `!== undefined` checks guard array index
      //   accesses or `Record<K, V>` lookups. tsconfig deliberately leaves
      //   `noUncheckedIndexedAccess: false` (the project predates that flag
      //   and turning it on would cascade dozens of new errors); the
      //   defensive runtime checks are still load-bearing — `parts[0]` IS
      //   undefined at runtime when the input doesn't split into 3 chunks,
      //   even though TS types it as `number`. SonarJS uses TS's view and
      //   thinks the check is dead.
      // - `argument-type`: misfires on `Array.prototype.includes(s)` when
      //   the array element type is narrower than `string` but
      //   `Object.keys()`-style inference erases the narrowing back to
      //   `string`. Same root cause — TS / SonarJS type-view mismatch.
      'sonarjs/no-useless-intersection': 'off',
      'sonarjs/function-return-type': 'off',
      'sonarjs/null-dereference': 'off',
      'sonarjs/different-types-comparison': 'off',
      'sonarjs/argument-type': 'off',
    },
  },
];
