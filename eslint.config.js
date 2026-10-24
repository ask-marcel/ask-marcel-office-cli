import { baseConfig } from './eslint.base.js';
import { CORE_DYNAMIC_IMPORT_BANS, CORE_FILES, CORE_IMPORT_BANS, CORE_READ_ONLY_FILES, READ_ONLY_IMPORT_BANS, STDOUT_BAN } from './eslint.ownership.js';

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
    // except the CLI writers (the CLI builder and the presenter's `output.ts`),
    // so a new MCP file is covered the day it lands, whatever its name. The
    // token helper's one JSON line moved to auth/ with its allowance (phase 4).
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**', 'src/composition/cli-shell.ts', 'src/presenter/output.ts'],
    rules: {
      'no-restricted-properties': ['error', STDOUT_BAN],
    },
  },
  {
    // Read-only code never holds the write graph (package split, D9): the read
    // files import neither write-graph.ts (its type included, so no read command
    // can be typed against it) nor a value from graph-client.ts, which builds the
    // client of the composition root: a read command runs on the graph that the
    // registry gives it. Type imports from graph-client.ts (`GraphError`) are
    // erased at build time.
    // The composition root (build-deps.ts) and graph-client.ts itself are the
    // exceptions. The write commands and their write-only helpers are in write/
    // since phase 5, outside this block. read/eslint.config.js applies the same
    // scope to read's files after the move.
    // scripts/check-read-bundle.ts checks the same boundary on the built code.
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**', 'src/composition/build-deps.ts', 'src/infra/graph-client.ts'],
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

// The ownership lists stay importable from here, where the checks read them from step 15 to phase 6.
export { CORE_FILES, CORE_READ_ONLY_FILES, SINGLE_PACKAGE_FILES } from './eslint.ownership.js';
