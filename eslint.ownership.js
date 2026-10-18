/*
 * The data the lint configs and the package checks share (package split): who
 * owns each module, and the import and stdout bans on those files. Pure data,
 * so a folder's eslint.config.js can import it without loading the root config
 * (two configs that each call baseConfig leave typescript-eslint with two
 * tsconfig roots). The root eslint.config.js and the folder configs apply
 * the bans; scripts/check-package-imports.ts reads the lists.
 */

/*
 * Import bans shared by more than one block of the lint configs. ESLint replaces a rule's
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

// The auth ladder (auth/src/infra/auth.ts; auth/eslint.config.js applies the
// ban) never loads the browser. An unattended session runs on it alone, and the
// auth `token` entry stands on it and must start fast. Browser rungs reach it
// only through the BrowserRungs port that auth-browser.ts supplies. Type
// imports are erased at build time, so they stay allowed.
const AUTH_LADDER_IMPORT_BANS = [
  {
    regex: '(^|/)(auth-browser|browser-auth|playwright-loader)(\\.[cm]?[jt]s)?$|^playwright(-core)?(/.*)?$',
    allowTypeImports: true,
    message: 'auth.ts is the browser-free half of auth: take a browser rung through the BrowserRungs port that auth-browser.ts supplies, never import the browser.',
  },
];

// The token helper (auth/src/token.ts, auth/src/composition/token-helper.ts;
// auth/eslint.config.js applies the ban) stays near bare Node startup: read and
// write spawn it for every bearer they lack. It never loads the CLI or the MCP
// server (the root's builders, auth's builders, the MCP SDK, zod), winston, the
// update notifier, the renderer or, statically, the browser; the browser
// ladder comes in by a dynamic import only for a person at a terminal. Type
// imports are erased, so they stay allowed. src/token.ts is a bundle entry, so
// it owns the one crash catch of its bundle (rule 17), and that catch prints
// only the error's type: a message can quote a token (LESSONS, 2026-10-04).
const TOKEN_HELPER_IMPORT_BANS = [
  {
    regex:
      '^(commander|winston|update-notifier|zod|@modelcontextprotocol/[^/]+|playwright(-core)?)(/.*)?$|(^|/)(cli|cli-shell|build-deps|mcp|mcp-shell|run-registry-command|auth-cli|auth-mcp|auth-program|auth-server|build-auth-deps)(\\.[cm]?[jt]s)?$|(^|/)presenter/|(^|/)infra/(logger|auth-browser|browser-auth|playwright-loader)(\\.[cm]?[jt]s)?$',
    allowTypeImports: true,
    message:
      'The token helper loads only the auth ladder, the file system and the token endpoint call; bring the browser in with a dynamic import, never the CLI or the MCP server.',
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
 * "What goes where"). The lint blocks read these lists, and
 * scripts/check-package-imports.ts reads them too: a write module may import only
 * core and write modules, and no read command and no core module may reach a
 * write module. A file joins a list when it becomes core or write.
 *
 * WRITE_FILES: the write commands and their write-only helpers. A new write
 * command joins the list (meta.test.ts pins the write set). Until the write
 * files move to write/ (package split, phase 5), the list also holds the write
 * package's own wiring under the root src/: its registry, its lifecycle entries
 * and its composition.
 *
 * A test file belongs to one package too (step 16): the check reads
 * SINGLE_PACKAGE_FILES for that, and no lint block reads it yet. A module in no
 * list is read's. A file under core/ or auth/ belongs to its folder's package
 * (phases 3 and 4): auth has no list.
 */
const WRITE_FILES = [
  'src/use-cases/commands/{create-mail-draft,create-reply-draft,create-forward-draft,update-mail-draft}.ts',
  'src/use-cases/commands/convert-{mail,calendar-event,group-post}-attachment-to-pdf.ts',
  'src/use-cases/commands/{draft-comment-splicer,draft-response,parse-recipients}.ts',
  'src/use-cases/commands/{write-commands,write-lifecycle-entries}.ts',
  'src/composition/{write-package-names,write-package,write-program,write-server,write-token-source,write-client,build-write-deps}.ts',
];

// CORE_FILES: the core modules that the read-only ban does not reach.
const CORE_FILES = [
  'src/composition/{cli-shell,mcp-shell,run-registry-command,date-zone,skeleton-cli,token-source}.ts',
  'src/use-cases/commands/{command-types,command-graph}.ts',
  'src/presenter/*.ts',
  'src/use-cases/ports/{filesystem,logger,process-runner,token-source}.ts',
  'src/infra/{write-graph,network-error,logger,auth-paths}.ts',
  'src/infra/{env-token-source,helper-token-source,token-helper-run,token-helper-answer,token-helper-locator,token-source-needed-by}.ts',
  'src/infra/{filesystem-bun,filesystem-node,process-runner-bun,process-runner-node,child-process-watch,rename-with-retry}.ts',
  'src/domain/{result,access-token,tenant-id,env-var,iso-datetime,json,page-range,jwt-utils,closest-names,token-fingerprint,teams-region}.ts',
  'src/domain/utilities/{base64,format-error,spo-tenant}.ts',
  // The shared test fakes (the plan's core/src/test-helpers).
  'src/test-helpers/{atomic-file-writes-contract,command-package-fake,fetch-mock,file-exists-contract,filesystem-fake,format-error,jwt,logger-fake}.ts',
  'src/test-helpers/{process-runner-contract,process-runner-fake,read-graph-fake,write-graph-fake,graph-client-fake}.ts',
];

// CORE_READ_ONLY_FILES: the core modules that the read-only ban also guards.
const CORE_READ_ONLY_FILES = [
  'src/use-cases/commands/{command-package,command-effect,read-only-post,build-command,format-zod-error,reject-unknown-params,resolve-command,tenant-option,odata-query}.ts',
  'src/use-cases/commands/{command-directory,bin-names,unknown-command}.ts',
  'src/use-cases/commands/{docs-builders,docs-render,first-sentence,output-path,office-extensions,date-zone}.ts',
  'src/use-cases/commands/{fetch-raw-bytes,quote-boundary,share-token,text-passthrough}.ts',
  'src/infra/{graph-request,read-graph,graph-substrate}.ts',
];

// SINGLE_PACKAGE_FILES: the wiring of the single package, which holds the three
// packages in one bin: its registry and lifecycle entries, its CLI, MCP server
// and composition root, its entries, and the full client over a token source.
// No package of the split takes them; phases 4 to 6 replace them with each
// package's own wiring.
const SINGLE_PACKAGE_FILES = [
  'src/{index,main}.ts',
  'src/composition/{cli,mcp,single-package,single-package-names,build-deps}.ts',
  'src/use-cases/commands/{index,lifecycle-entries}.ts',
  'src/infra/graph-client.ts',
];

// The option of the stdout ban, which the root eslint.config.js and the core, auth and write configs
// apply to their files too (the MCP gateway moved to core in phase 3, the auth MCP tools to auth in
// phase 4, and the write MCP server moves to write in phase 5).
const STDOUT_BAN = {
  object: 'process',
  property: 'stdout',
  message:
    'stdout is the MCP JSON-RPC frame channel — writing to it corrupts the protocol and no test can see it. Return text from the tool handler (renderToString / renderErrorToString) instead. Use the stderr-backed Logger port if you need diagnostics. Only the CLI writers (cli-shell.ts, presenter/output.ts, token-helper.ts) own stdout.',
};

export {
  AUTH_LADDER_IMPORT_BANS,
  CORE_DYNAMIC_IMPORT_BANS,
  CORE_FILES,
  CORE_IMPORT_BANS,
  CORE_READ_ONLY_FILES,
  READ_ONLY_IMPORT_BANS,
  SINGLE_PACKAGE_FILES,
  STDOUT_BAN,
  TOKEN_HELPER_IMPORT_BANS,
  WRITE_FILES,
};
