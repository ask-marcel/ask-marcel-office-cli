import { baseConfig } from './eslint.base.js';
import { CORE_DYNAMIC_IMPORT_BANS, CORE_FILES, CORE_IMPORT_BANS, CORE_READ_ONLY_FILES, READ_ONLY_IMPORT_BANS, STDOUT_BAN, WRITE_FILES } from './eslint.ownership.js';

/** @type {import('eslint').Linter.Config[]} */
export default [
  ...baseConfig(import.meta.dirname),
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
      'no-restricted-properties': ['error', STDOUT_BAN],
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
    // The write commands and their write-only helpers (WRITE_FILES), and the
    // registry and command framework that hold both kinds are the exceptions.
    // scripts/check-read-bundle.ts checks the same boundary on the built code.
    files: ['src/use-cases/commands/**/*.ts', 'src/infra/read-graph.ts', 'src/infra/graph-request.ts', 'src/infra/graph-substrate.ts'],
    ignores: ['**/*.test.ts', ...WRITE_FILES, 'src/use-cases/commands/{index,command-graph,command-types}.ts'],
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
    // A file joins CORE_FILES or CORE_READ_ONLY_FILES when it becomes core.
    // Their globs name `src/...`: here the root's files, and in
    // core/eslint.config.js, which applies the same two blocks, core's own.
    files: CORE_FILES,
    ignores: ['**/*.test.ts'],
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: CORE_IMPORT_BANS }],
      'no-restricted-syntax': ['error', ...CORE_DYNAMIC_IMPORT_BANS],
    },
  },
  {
    // The core files that the read-only block above also guards keep both bans.
    files: CORE_READ_ONLY_FILES,
    rules: {
      '@typescript-eslint/no-restricted-imports': ['error', { patterns: [...READ_ONLY_IMPORT_BANS, ...CORE_IMPORT_BANS] }],
      'no-restricted-syntax': ['error', ...CORE_DYNAMIC_IMPORT_BANS],
    },
  },
  {
    // Each workspace folder lints itself with its own eslint.config.js (D7).
    ignores: ['dist/**', '.stryker-tmp/**', 'reports/**', 'docs/**', 'scripts/**', '.claude/**', '.agents/**', 'auth/**', 'core/**', 'read/**', 'write/**'],
  },
];

// The ownership lists stay importable from here, where the checks have read them since step 15.
export { AUTH_FILES, CORE_FILES, CORE_READ_ONLY_FILES, SINGLE_PACKAGE_FILES, WRITE_FILES } from './eslint.ownership.js';
