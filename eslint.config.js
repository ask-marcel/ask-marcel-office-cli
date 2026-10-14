import pluginJs from '@eslint/js';
import prettier from 'eslint-plugin-prettier';
import securityPlugin from 'eslint-plugin-security';
import sonarjsPlugin from 'eslint-plugin-sonarjs';
import unicornPlugin from 'eslint-plugin-unicorn';
import globals from 'globals';
import tsPlugin from 'typescript-eslint';

/*
 * Import bans shared by more than one block below. ESLint replaces a rule's
 * options when two blocks match the same file (it does not merge them), so a
 * file under two bans gets a block that lists both.
 */
const READ_ONLY_IMPORT_BANS = [
  {
    regex: '(^|/)write-graph(\\.[cm]?[jt]s)?$',
    message: 'Read-only code never holds the write graph: type a read command against ReadGraph (read-graph.ts), which POSTs only to the two query endpoints.',
  },
  {
    regex: '(^|/)graph-client(\\.[cm]?[jt]s)?$',
    allowTypeImports: true,
    message: 'graph-client.ts builds the full client, write graph included: read code takes the read graph (read-graph.ts) instead.',
  },
];

const CORE_IMPORT_BANS = [
  {
    regex: '^(\\./|\\.\\./(.*/)?commands/)(index|login|logout|status|login-status|token|token-args|token-tier-capability|lifecycle-entries)(\\.[cm]?[jt]s)?$',
    message: 'Core serves any package: take the registry and the lifecycle entries from the CommandPackage the composition root passes in, never import them.',
  },
  {
    // A folder import loads the folder's index.ts: commands/ is the registry, and
    // a path of dots only (`.`, `..`, `../..`) can reach src/index.ts.
    regex: '(^|/)commands/?$|^[./]+$',
    message: 'A folder import loads its index.ts: commands/ is the registry and src/ is the package entry. Take the registry from the CommandPackage instead.',
  },
  {
    regex: '^(\\.\\./)+(index|main|token)(\\.[cm]?[jt]s)?$|^ask-marcel-office-cli(/.*)?$',
    message: 'Core never imports a package entry: index.ts re-exports the registry and buildDeps, and main.ts and token.ts are bundle entries.',
  },
  {
    regex: '^\\.\\.?/(.*/)?(single-package|single-package-names|build-deps|cli|mcp|auth-cli|auth-mcp)(\\.[cm]?[jt]s)?$',
    message: 'Core never reaches the composition root that wires the single package: take what you need as a parameter (CommandPackage, a hook) instead.',
  },
];

// no-restricted-imports does not see `import(...)`, so the same specifiers are
// banned in a dynamic import too. Other dynamic imports stay legal in core (the
// Bun file system loads its Node half that way).
const CORE_DYNAMIC_IMPORT_BANS = CORE_IMPORT_BANS.map(({ regex, message }) => ({
  selector: `ImportExpression[source.value=/${regex.replaceAll('/', '\\/')}/]`,
  message,
}));

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
    // An allowlist since package-split step 13: stdout is banned everywhere
    // except the CLI writers (the CLI builder, the presenter's `output.ts`, and
    // the token helper's one JSON line), so a new MCP file is covered the day it
    // lands, whatever its name.
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**', 'src/composition/cli-shell.ts', 'src/presenter/output.ts', 'src/composition/token-helper.ts'],
    rules: {
      'no-restricted-properties': [
        'error',
        {
          object: 'process',
          property: 'stdout',
          message:
            'stdout is the MCP JSON-RPC frame channel — writing to it corrupts the protocol and no test can see it. Return text from the tool handler (renderToString / renderErrorToString) instead. Use the stderr-backed Logger port if you need diagnostics. Only the CLI writers (cli-shell.ts, presenter/output.ts, token-helper.ts) own stdout.',
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
    // allowed. src/token.ts is a bundle entry like src/main.ts, so it owns the
    // one crash catch of its bundle (rule 17), and that catch prints only the
    // error's type: a message can quote a token (LESSONS, 2026-10-04).
    files: ['src/token.ts', 'src/composition/token-helper.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              regex:
                '^(commander|winston|update-notifier|playwright(-core)?)(/.*)?$|(^|/)(cli|cli-shell|build-deps|mcp|mcp-shell|run-registry-command)(\\.[cm]?[jt]s)?$|(^|/)presenter/|(^|/)infra/(logger|auth-browser|browser-auth|playwright-loader)(\\.[cm]?[jt]s)?$',
              allowTypeImports: true,
              message: 'The token helper loads only the auth ladder, the file system and the token endpoint call; bring the browser in with a dynamic import, never the CLI.',
            },
          ],
        },
      ],
    },
  },
  {
    // Read-only code never holds the write graph (package split, D9): the read
    // commands, their helpers and the read graph's own modules import neither
    // write-graph.ts (its type included, so no read command can be typed against
    // it) nor a value from graph-client.ts, which builds the full client. Type
    // imports from graph-client.ts (`GraphError`) are erased at build time.
    // The write commands, their write-only helpers, and the registry and
    // command framework that hold both kinds are the exceptions. A new write
    // command joins the list (meta.test.ts pins the write set).
    // scripts/check-read-bundle.ts checks the same boundary on the built code.
    files: ['src/use-cases/commands/**/*.ts', 'src/infra/read-graph.ts', 'src/infra/graph-request.ts', 'src/infra/graph-substrate.ts'],
    ignores: [
      '**/*.test.ts',
      'src/use-cases/commands/{create-mail-draft,create-reply-draft,create-forward-draft,update-mail-draft}.ts',
      'src/use-cases/commands/convert-{mail,calendar-event,group-post}-attachment-to-pdf.ts',
      'src/use-cases/commands/{draft-comment-splicer,draft-response,parse-recipients}.ts',
      'src/use-cases/commands/{index,command-graph,command-types}.ts',
    ],
    rules: {
      '@typescript-eslint/no-restricted-imports': [
        'error',
        {
          patterns: READ_ONLY_IMPORT_BANS,
        },
      ],
    },
  },
  {
    // Core never knows which package it serves (package split, step 13). The
    // two blocks below hold the plan's core list ("What goes where"): the CLI
    // builder, the MCP gateway, run-registry-command, the docs builders, the
    // command framework, the presenter, the Graph request core and both graphs,
    // the token source, the file system, process runner and logger with their
    // ports, and the shared domain types. They take the registry, the lifecycle
    // entries and the names from the `CommandPackage` the composition root
    // hands them, so they never import the registry (by path, by folder, through
    // the package entry, or with a dynamic import), a lifecycle command, or the
    // modules that wire the single package (`single-package.ts`, `build-deps.ts`,
    // `cli.ts`, `mcp.ts`, `auth-cli.ts`, `auth-mcp.ts`, `single-package-names.ts`).
    // A file joins a list when it becomes core (step 15 extracts the shared helpers).
    files: [
      'src/composition/{cli-shell,mcp-shell,run-registry-command,date-zone}.ts',
      'src/use-cases/commands/{command-types,command-graph}.ts',
      'src/presenter/*.ts',
      'src/use-cases/ports/{filesystem,logger,process-runner,token-source}.ts',
      'src/infra/{write-graph,network-error,logger}.ts',
      'src/infra/{env-token-source,helper-token-source,token-helper-run,token-helper-answer,token-helper-locator}.ts',
      'src/infra/{filesystem-bun,filesystem-node,process-runner-bun,process-runner-node,child-process-watch,rename-with-retry}.ts',
      'src/domain/{result,access-token,tenant-id,env-var,iso-datetime,json,page-range,jwt-utils,closest-names,token-fingerprint,teams-region}.ts',
      'src/domain/utilities/base64.ts',
    ],
    ignores: ['**/*.test.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: CORE_IMPORT_BANS }],
      'no-restricted-syntax': ['error', ...CORE_DYNAMIC_IMPORT_BANS],
    },
  },
  {
    // The core files that the read-only block above also guards keep both bans.
    files: [
      'src/use-cases/commands/{command-package,command-effect,read-only-post,build-command,format-zod-error,reject-unknown-params,resolve-command}.ts',
      'src/use-cases/commands/{command-directory,bin-names,unknown-command}.ts',
      'src/use-cases/commands/{docs-builders,docs-render,first-sentence,output-path,office-extensions,date-zone}.ts',
      'src/infra/{graph-request,read-graph,graph-substrate}.ts',
    ],
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: [...READ_ONLY_IMPORT_BANS, ...CORE_IMPORT_BANS] }],
      'no-restricted-syntax': ['error', ...CORE_DYNAMIC_IMPORT_BANS],
    },
  },
  {
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**', 'docs/**', 'scripts/**', '.claude/**', '.agents/**'],
  },
];
