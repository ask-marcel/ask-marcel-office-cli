import pluginJs from '@eslint/js';
import prettier from 'eslint-plugin-prettier';
import securityPlugin from 'eslint-plugin-security';
import sonarjsPlugin from 'eslint-plugin-sonarjs';
import unicornPlugin from 'eslint-plugin-unicorn';
import globals from 'globals';
import tsPlugin from 'typescript-eslint';

/** @type {import('eslint').Linter.Config[]} */
export default [
  pluginJs.configs.recommended,
  ...tsPlugin.configs.recommended,
  securityPlugin.configs.recommended,
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
      'no-restricted-imports': [
        'error',
        {
          paths: [
            {
              name: 'bun:test',
              importNames: ['mock'],
              message:
                '`mock` from bun:test is forbidden — it leaks across test files. Use dependency injection: refactor the production code to accept the SDK as a parameter, then pass a fake at construction.',
            },
          ],
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
  ...(process.env['LINT_STRICT']
    ? [
        {
          files: ['src/**/*.ts'],
          languageOptions: {
            parserOptions: {
              projectService: true,
              tsconfigRootDir: import.meta.dirname,
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
  {
    // The MCP gateway serves JSON-RPC frames over stdout. ANY other write to
    // that stream corrupts the protocol, and nothing else can catch it:
    // `mcp.test.ts` drives InMemoryTransport (never touches stdout), and even a
    // real stdio handshake stays green because the SDK's client-side ReadBuffer
    // silently skips lines it cannot parse (verified 2026-07-17 by injecting a
    // banner — the Client-based probe reported a clean 5-tool handshake against
    // a provably corrupt stream).
    //
    // So this rule is the enforcement. Render via `renderToString` /
    // `renderErrorToString` (pure) and return the string; never write it.
    // Scoped to mcp.ts deliberately: `cli.ts` and `presenter/output.ts` own
    // stdout legitimately in CLI mode.
    files: ['src/composition/mcp.ts'],
    rules: {
      'no-restricted-properties': [
        'error',
        {
          object: 'process',
          property: 'stdout',
          message:
            'stdout is the MCP JSON-RPC frame channel — writing to it corrupts the protocol and no test can see it. Return text from the tool handler (renderToString / renderErrorToString) instead. Use the stderr-backed Logger port if you need diagnostics.',
        },
      ],
    },
  },
  {
    // The auth ladder (token cache, refresh, status) never loads the browser.
    // An unattended session runs on it alone, and the auth package's `token`
    // entry will stand on it and must start fast. Browser rungs reach it only
    // through the BrowserRungs port that auth-browser.ts supplies. Type imports
    // are erased at build time, so they stay allowed.
    files: ['src/infra/auth.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              regex: '(^|/)(auth-browser|browser-auth|playwright-loader)(\\.[cm]?[jt]s)?$|^playwright(-core)?(/.*)?$',
              allowTypeImports: true,
              message: 'auth.ts is the browser-free half of auth: take a browser rung through the BrowserRungs port that auth-browser.ts supplies, never import the browser.',
            },
          ],
        },
      ],
    },
  },
  {
    // The `token` helper starts on its own entry (dist/token.js) and must stay
    // near bare Node startup: read and write spawn it for every bearer they lack.
    // It never loads the CLI, winston, the update notifier, the renderer or,
    // statically, the browser; the browser ladder comes in by a dynamic import
    // only for a person at a terminal. Type imports are erased, so they stay
    // allowed.
    files: ['src/token.ts', 'src/composition/token-helper.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              regex:
                '^(commander|winston|update-notifier|playwright(-core)?)(/.*)?$|(^|/)(cli|build-deps|mcp|run-registry-command)(\\.[cm]?[jt]s)?$|(^|/)presenter/|(^|/)infra/(logger|auth-browser|browser-auth|playwright-loader)(\\.[cm]?[jt]s)?$',
              allowTypeImports: true,
              message: 'The token helper loads only the auth ladder, the file system and the token endpoint call; bring the browser in with a dynamic import, never the CLI.',
            },
          ],
        },
      ],
    },
  },
  {
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**', 'docs/**', 'scripts/**', '.claude/**', '.agents/**'],
  },
];
