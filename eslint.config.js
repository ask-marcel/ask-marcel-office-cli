import { baseConfig } from './eslint.base.js';

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

/*
 * Who owns each module after the package split (docs/plans/2026-10-01-package-split.md,
 * "What goes where"). The blocks below read these lists, and
 * scripts/check-package-imports.ts reads them too: a write module may import only
 * core and write modules, and no read command and no core module may reach a
 * write module. A file joins a list when it becomes core or write.
 *
 * WRITE_FILES: the write commands and their write-only helpers. A new write
 * command joins the list (meta.test.ts pins the write set).
 *
 * A test file belongs to one package too (step 16): the check reads AUTH_FILES
 * and SINGLE_PACKAGE_FILES for that, and no lint block reads them yet. A module
 * in no list is read's.
 */
const WRITE_FILES = [
  'src/use-cases/commands/{create-mail-draft,create-reply-draft,create-forward-draft,update-mail-draft}.ts',
  'src/use-cases/commands/convert-{mail,calendar-event,group-post}-attachment-to-pdf.ts',
  'src/use-cases/commands/{draft-comment-splicer,draft-response,parse-recipients}.ts',
];

// CORE_FILES: the core modules that the read-only ban does not reach.
const CORE_FILES = [
  'src/composition/{cli-shell,mcp-shell,run-registry-command,date-zone}.ts',
  'src/use-cases/commands/{command-types,command-graph}.ts',
  'src/presenter/*.ts',
  'src/use-cases/ports/{filesystem,logger,process-runner,token-source}.ts',
  'src/infra/{write-graph,network-error,logger}.ts',
  'src/infra/{env-token-source,helper-token-source,token-helper-run,token-helper-answer,token-helper-locator,token-source-needed-by}.ts',
  'src/infra/{filesystem-bun,filesystem-node,process-runner-bun,process-runner-node,child-process-watch,rename-with-retry}.ts',
  'src/domain/{result,access-token,tenant-id,env-var,iso-datetime,json,page-range,jwt-utils,closest-names,token-fingerprint,teams-region}.ts',
  'src/domain/utilities/{base64,spo-tenant}.ts',
  // The shared test fakes (the plan's core/src/test-helpers).
  'src/test-helpers/{atomic-file-writes-contract,command-package-fake,fetch-mock,file-exists-contract,filesystem-fake,format-error,jwt,logger-fake}.ts',
  'src/test-helpers/{process-runner-contract,process-runner-fake,read-graph-fake,write-graph-fake}.ts',
];

// CORE_READ_ONLY_FILES: the core modules that the read-only ban also guards.
const CORE_READ_ONLY_FILES = [
  'src/use-cases/commands/{command-package,command-effect,read-only-post,build-command,format-zod-error,reject-unknown-params,resolve-command}.ts',
  'src/use-cases/commands/{command-directory,bin-names,unknown-command}.ts',
  'src/use-cases/commands/{docs-builders,docs-render,first-sentence,output-path,office-extensions,date-zone}.ts',
  'src/use-cases/commands/{fetch-raw-bytes,quote-boundary,share-token,text-passthrough}.ts',
  'src/infra/{graph-request,read-graph,graph-substrate}.ts',
];

// AUTH_FILES: the auth commands of the command directory (login, logout, status,
// token) with their helpers, the auth ladder, the browser sign-in and the token
// helper's entry.
const AUTH_FILES = [
  'src/use-cases/commands/{login,logout,status,login-status,token,token-args,token-tier-capability}.ts',
  'src/use-cases/ports/{token-issuer,token-report}.ts',
  'src/composition/{auth-cli,auth-mcp,token-helper}.ts',
  'src/infra/{auth,auth-browser,auth-paths,browser-auth,playwright-loader,token-cache-lock}.ts',
  'src/token.ts',
];

// SINGLE_PACKAGE_FILES: the wiring of the single package, which holds the three
// packages in one bin: its registry and lifecycle entries, its CLI, MCP server
// and composition root, its entries, the full client over the in-process ladder,
// and the fakes of that client and that ladder. No package of the split takes
// them; phases 4 to 6 replace them with each package's own wiring.
const SINGLE_PACKAGE_FILES = [
  'src/{index,main}.ts',
  'src/composition/{cli,mcp,single-package,single-package-names,build-deps}.ts',
  'src/use-cases/commands/{index,lifecycle-entries}.ts',
  'src/infra/{graph-client,auth-token-source}.ts',
  'src/test-helpers/{graph-client-fake,auth-manager-fake}.ts',
];

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

export { AUTH_FILES, CORE_FILES, CORE_READ_ONLY_FILES, SINGLE_PACKAGE_FILES, WRITE_FILES };
