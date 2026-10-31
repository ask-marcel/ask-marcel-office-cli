#!/usr/bin/env bun
/*
 * Violation fixtures for the ESLint rules that apply to some paths only.
 *
 * A path-scoped rule stops applying the moment the files it names move, and
 * nothing reports it: lint just gets quieter. The package split moves every
 * file out of `src/`, so each such rule gets a fixture here that must be
 * rejected, plus a control proving the rule is what rejects it (rule 15.10: a
 * gate only ever seen green is a hypothesis).
 *
 * Fixtures are linted from memory through `--stdin-filename`; nothing is
 * written. Each fixture names the real file its rule guards, and that file
 * must exist: when it moves, the fixture fails until both the fixture path and
 * the rule's glob in eslint.config.js move with it. (A virtual path alone would
 * keep matching the old glob and stay green while the moved file lost the rule.)
 * One move is followed: a root file found under core/, auth/, write/ or read/
 * instead (package split, phases 3 to 6) is linted from that folder with its
 * config, which must hold the rule. A type-aware fixture follows its file too,
 * and then proves that folder's lint:strict. A path with no cwd that exists in
 * two folders names no one file, so its fixture fails.
 *
 * The strict fixtures run through `bun run lint:strict`, the command CI runs,
 * so they prove that script still turns the type-aware block on, not only that
 * the config holds it. The plain fixtures run eslint with LINT_STRICT unset.
 *
 * Exit codes:
 *   0  every fixture fires its rule where it must, and stays silent where it must
 *   1  a rule did not fire on its fixture, fired on its control, a guarded file
 *      is missing, or eslint did not lint the text (parse error, ignored path,
 *      unreadable output)
 */

type Fixture = {
  readonly name: string;
  readonly path: string;
  readonly source: string;
  readonly rule: string;
  readonly fires: boolean;
  readonly strict: boolean;
  // A workspace folder (package split, D7): the fixture is linted from there,
  // with that folder's eslint.config.js and lint:strict script, and its path is
  // relative to it.
  readonly cwd?: string;
};

const STDOUT_WRITE = "export const f = (): void => {\n  process.stdout.write('banner');\n};\n";
const NEEDLESS_ASSERTION = 'const n = 1;\nexport const f = (): number => n as number;\n';
const STRING_REJECTION = "export const f = (): Promise<never> => Promise.reject('nope');\n";
const BROWSER_IMPORT = "import { createBrowserAuth } from './browser-auth.ts';\nexport const f = (): unknown => createBrowserAuth;\n";
const WINSTON_IMPORT = "import winston from 'winston';\nexport const f = (): unknown => winston;\n";
const AUTH_SERVER_IMPORT = "import { buildAuthMcpServer } from './auth-server.ts';\nexport const f = (): unknown => buildAuthMcpServer;\n";
const AUTH_PROGRAM_IMPORT = "import { buildAuthCli } from './auth-program.ts';\nexport const f = (): unknown => buildAuthCli;\n";
const MCP_TYPES_IMPORT = "import { CallToolResultSchema } from '@modelcontextprotocol/sdk/types.js';\nexport const f = (): unknown => CallToolResultSchema;\n";
const ZOD_IMPORT = "import { z } from 'zod';\nexport const f = (): unknown => z;\n";
const WRITE_GRAPH_IMPORT = "import { createWriteGraph } from '../../infra/write-graph.ts';\nexport const f = (): unknown => createWriteGraph;\n";
const WRITE_GRAPH_TYPE_IMPORT = "import type { WriteGraph } from '../../infra/write-graph.ts';\nexport type F = WriteGraph;\n";
const ROOT_CLIENT_IMPORT = "import { createGraphClient } from '../../infra/graph-client.ts';\nexport const f = (): unknown => createGraphClient;\n";
const CORE_WRITE_GRAPH_IMPORT = "import { createWriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';\nexport const f = (): unknown => createWriteGraph;\n";
const CORE_WRITE_GRAPH_TYPE_IMPORT = "import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';\nexport type F = WriteGraph;\n";
const COMPOSITION_CLIENT_IMPORT = "import { createTokenSourceGraphClient } from '../infra/graph-client.ts';\nexport const f = (): unknown => createTokenSourceGraphClient;\n";
const SIBLING_WRITE_GRAPH_IMPORT = "import { createWriteGraph } from './write-graph.ts';\nexport const f = (): unknown => createWriteGraph;\n";
const REGISTRY_IMPORT = "import { commands } from '../use-cases/commands/index.ts';\nexport const f = (): unknown => commands;\n";
const SIBLING_REGISTRY_IMPORT = "import { commands } from './index.ts';\nexport const f = (): unknown => commands;\n";
const LIFECYCLE_ENTRIES_IMPORT = "import { LIFECYCLE_ENTRIES } from './lifecycle-entries.ts';\nexport const f = (): unknown => LIFECYCLE_ENTRIES;\n";
const LOGIN_IMPORT = "import * as login from '../use-cases/commands/login.ts';\nexport const f = (): unknown => login;\n";
const SINGLE_PACKAGE_IMPORT = "import { commandPackage } from './single-package.ts';\nexport const f = (): unknown => commandPackage;\n";
const SINGLE_PACKAGE_NAMES_IMPORT = "import { SINGLE_PACKAGE_NAMES } from '../../composition/single-package-names.ts';\nexport const f = (): unknown => SINGLE_PACKAGE_NAMES;\n";
const BARREL_REGISTRY_IMPORT = "import { commands } from '../index.ts';\nexport const f = (): unknown => commands;\n";
const REGISTRY_DIRECTORY_IMPORT = "import * as registry from '../use-cases/commands';\nexport const f = (): unknown => registry;\n";
const SIBLING_DIRECTORY_IMPORT = "import * as registry from '.';\nexport const f = (): unknown => registry;\n";
const SELF_PACKAGE_IMPORT = "import { commands } from 'ask-marcel-office-cli';\nexport const f = (): unknown => commands;\n";
const DYNAMIC_REGISTRY_IMPORT = "export const f = async (): Promise<unknown> => import('../use-cases/commands/index.ts');\n";
const DYNAMIC_SIBLING_IMPORT = "export const f = async (): Promise<unknown> => import('./filesystem-node.ts');\n";
const CORE_FAKE_IMPORT = "import { fakeLogger } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';\nexport const f = (): unknown => fakeLogger;\n";
const SIBLING_FAKE_IMPORT = "import { fakeLogger } from '../test-helpers/logger-fake.ts';\nexport const f = (): unknown => fakeLogger;\n";
// Rule 15 (no inline ignores) and rule 13 (no mocks), from eslint.base.js, so
// every folder holds them.
const INLINE_DISABLE = "// eslint-disable-next-line no-console\nexport const f = (): void => {\n  console.log('hidden');\n};\n";
const TS_EXPECT_ERROR = "// @ts-expect-error: a described one is banned too\nexport const n: number = 'a';\n";
const TOOL_MARKER = '// prettier-ignore\nexport const n = 1;\n';
const PLAIN_COMMENT = '// A plain comment names no tool.\nexport const n = 1;\n';
const SPY_IMPORT = "import { spyOn } from 'bun:test';\nexport const f = (): unknown => spyOn;\n";
const TEST_IMPORT = "import { expect, test } from 'bun:test';\ntest('t', () => {\n  expect(1).toBe(1);\n});\n";
const CALL_ASSERTION = "import { expect, test } from 'bun:test';\ntest('t', () => {\n  const fake = { calls: 0 };\n  expect(fake).toHaveBeenCalledWith(1);\n});\n";
const CLASS_DECLARATION = 'export class Box {\n  readonly n = 1;\n}\n';
const INLINE_TYPE_IMPORT = "import { ok, type Result } from './result.ts';\nexport const f = (): Result<number, string> => ok(1);\n";
const CURRIED_ARROW = 'export const add = (a: number) => (b: number): number => a + b;\n';
const CREATE_FACTORY = 'export const createAdder = (a: number) => (b: number): number => a + b;\n';
// Rule 35: eleven paths (ten ifs) is one over the limit, and nine ifs is at it.
const pathsFn = (ifs: number): string =>
  `export const f = (n: number): number => {\n${Array.from({ length: ifs }, (_, i) => `  if (n === ${i}) return ${i};\n`).join('')}  return -1;\n};\n`;
const ELEVEN_PATHS = pathsFn(10);
const TEN_PATHS = pathsFn(9);
const TRY_CATCH = "export const f = (raw: string): unknown => {\n  try {\n    return JSON.parse(raw);\n  } catch {\n    return undefined;\n  }\n};\n";
const CALL_ASSERTION_IN_SOURCE = "export const f = (e: { toHaveBeenCalled: () => void }): void => {\n  e.toHaveBeenCalled();\n};\n";

// With linterOptions.noInlineConfig, ESLint reports each directive comment
// with no rule id. The fixtures name that report with this label.
const INLINE_CONFIG = 'linterOptions.noInlineConfig';

// One real source file per workspace folder, relative to the folder.
const WORKSPACE_FILES: ReadonlyArray<readonly [string, string]> = [
  ['core', 'src/domain/result.ts'],
  ['auth', 'src/composition/auth-program.ts'],
  ['read', 'src/use-cases/commands/list-drives.ts'],
  ['write', 'src/composition/write-program.ts'],
];

// One real test file per workspace folder, relative to the folder.
const WORKSPACE_TESTS: ReadonlyArray<readonly [string, string]> = [
  ['core', 'src/composition/cli-shell-flags.test.ts'],
  ['auth', 'src/composition/auth-names.test.ts'],
  ['read', 'src/composition/build-deps.test.ts'],
  ['write', 'src/composition/cli-write.test.ts'],
];

// The rule 13 and rule 15 bans of eslint.base.js, in each folder.
const baseBanFixtures = ([cwd, path]: readonly [string, string], [, testPath]: readonly [string, string]): ReadonlyArray<Fixture> => [
  { name: `${cwd}: an inline eslint directive is reported`, path, source: INLINE_DISABLE, rule: INLINE_CONFIG, fires: true, strict: false, cwd },
  { name: `${cwd}: an inline eslint directive hides nothing`, path, source: INLINE_DISABLE, rule: 'no-console', fires: true, strict: false, cwd },
  { name: `${cwd}: a described @ts-expect-error is rejected`, path, source: TS_EXPECT_ERROR, rule: '@typescript-eslint/ban-ts-comment', fires: true, strict: false, cwd },
  { name: `${cwd}: the ignore marker of another tool is rejected`, path, source: TOOL_MARKER, rule: 'no-warning-comments', fires: true, strict: false, cwd },
  { name: `control: ${cwd}: a plain comment is not a marker`, path, source: PLAIN_COMMENT, rule: 'no-warning-comments', fires: false, strict: false, cwd },
  { name: `${cwd}: a test never imports spyOn`, path: testPath, source: SPY_IMPORT, rule: 'no-restricted-imports', fires: true, strict: false, cwd },
  { name: `control: ${cwd}: a test imports expect and test`, path: testPath, source: TEST_IMPORT, rule: 'no-restricted-imports', fires: false, strict: false, cwd },
  { name: `${cwd}: a test never asserts on recorded calls`, path: testPath, source: CALL_ASSERTION, rule: 'no-restricted-syntax', fires: true, strict: false, cwd },
  { name: `${cwd}: a class is rejected`, path, source: CLASS_DECLARATION, rule: 'no-restricted-syntax', fires: true, strict: false, cwd },
  { name: `${cwd}: an inline type specifier is rejected`, path, source: INLINE_TYPE_IMPORT, rule: 'no-restricted-syntax', fires: true, strict: false, cwd },
  { name: `${cwd}: a curried arrow is rejected`, path, source: CURRIED_ARROW, rule: 'no-restricted-syntax', fires: true, strict: false, cwd },
  { name: `control: ${cwd}: a createX factory may return a function`, path, source: CREATE_FACTORY, rule: 'no-restricted-syntax', fires: false, strict: false, cwd },
  { name: `${cwd}: a function with eleven paths is rejected`, path, source: ELEVEN_PATHS, rule: 'complexity', fires: true, strict: false, cwd },
  { name: `control: ${cwd}: a function with ten paths is kept`, path, source: TEN_PATHS, rule: 'complexity', fires: false, strict: false, cwd },
];

// Rule 17: a try in a use-case, one file of each folder. Core's own blocks
// replace the base's options, so core's file is one that a core block covers.
const USE_CASE_FILES: ReadonlyArray<readonly [string, string]> = [
  ['core', 'src/use-cases/commands/docs-builders.ts'],
  ['auth', 'src/use-cases/commands/status.ts'],
  ['read', 'src/use-cases/commands/list-drives.ts'],
  ['write', 'src/use-cases/commands/draft-response.ts'],
];

const tryBanFixtures = ([cwd, path]: readonly [string, string]): ReadonlyArray<Fixture> => [
  { name: `${cwd}: a use-case never catches`, path, source: TRY_CATCH, rule: 'no-restricted-syntax', fires: true, strict: false, cwd },
];

// Rule 37 (ADR 0004): the layer zones of eslint.base.js. A use-case of each
// folder never imports infra, by a relative path or by a core export path, and
// takes its types from a port: a converter adapter too, which reaches a read
// command through the Converters port. Core holds a file in every layer, so it
// proves the other zones.
const USE_CASE_INFRA = "import type { ReadGraph } from '../../infra/read-graph.ts';\nexport type F = ReadGraph;\n";
const USE_CASE_CORE_INFRA = "import type { ReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';\nexport type F = ReadGraph;\n";
const USE_CASE_PORT = "import type { ReadGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';\nexport type F = ReadGraph;\n";
const USE_CASE_ADAPTER = "import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';\nexport const f = (): unknown => htmlToMarkdown;\n";
const LAYER_PORT = "import type { Logger } from '../use-cases/ports/logger.ts';\nexport type F = Logger;\n";
const LAYER_USE_CASE = "import { execute } from '../use-cases/commands/login.ts';\nexport const f = (): unknown => execute;\n";
const LAYER_COMPOSITION = "import { main } from '../composition/main.ts';\nexport const f = (): unknown => main;\n";

const layerZoneFixtures = ([cwd, path]: readonly [string, string]): ReadonlyArray<Fixture> => [
  { name: `${cwd}: a use-case never imports infra`, path, source: USE_CASE_INFRA, rule: 'no-restricted-imports', fires: true, strict: false, cwd },
  { name: `${cwd}: a use-case never imports core's infra`, path, source: USE_CASE_CORE_INFRA, rule: 'no-restricted-imports', fires: true, strict: false, cwd },
  { name: `control: ${cwd}: a use-case takes a type from a port`, path, source: USE_CASE_PORT, rule: 'no-restricted-imports', fires: false, strict: false, cwd },
  { name: `${cwd}: a use-case never imports a converter adapter`, path, source: USE_CASE_ADAPTER, rule: 'no-restricted-imports', fires: true, strict: false, cwd },
];

const CORE_LAYER_FIXTURES: ReadonlyArray<Fixture> = [
  { name: 'core: domain never imports a port', path: 'src/domain/access-token.ts', source: LAYER_PORT, rule: 'no-restricted-imports', fires: true, strict: false, cwd: 'core' },
  { name: 'core: the presenter never imports a port', path: 'src/presenter/error-hints.ts', source: LAYER_PORT, rule: 'no-restricted-imports', fires: true, strict: false, cwd: 'core' },
  { name: 'core: infra never imports a use-case', path: 'src/infra/auth-paths.ts', source: LAYER_USE_CASE, rule: 'no-restricted-imports', fires: true, strict: false, cwd: 'core' },
  { name: 'core: infra never imports the composition root', path: 'src/infra/auth-paths.ts', source: LAYER_COMPOSITION, rule: 'no-restricted-imports', fires: true, strict: false, cwd: 'core' },
  { name: 'control: core: infra implements a port', path: 'src/infra/auth-paths.ts', source: LAYER_PORT, rule: 'no-restricted-imports', fires: false, strict: false, cwd: 'core' },
  { name: 'auth: infra never imports a use-case', path: 'src/infra/auth-browser.ts', source: LAYER_USE_CASE, rule: 'no-restricted-imports', fires: true, strict: false, cwd: 'auth' },
  { name: 'read: infra never imports a use-case', path: 'src/infra/eml-parser-adapter.ts', source: LAYER_USE_CASE, rule: 'no-restricted-imports', fires: true, strict: false, cwd: 'read' },
];

const FIXTURES: ReadonlyArray<Fixture> = [
  { name: 'the MCP server never writes stdout', path: 'src/composition/mcp.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
  { name: 'the MCP gateway never writes stdout', path: 'src/composition/mcp-shell.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
  { name: 'auth: the auth bin never writes stdout', path: 'src/main.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false, cwd: 'auth' },
  { name: 'the auth MCP tools never write stdout', path: 'src/composition/auth-mcp.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
  { name: 'write: the write bin never writes stdout', path: 'src/main.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false, cwd: 'write' },
  { name: 'read: the read bin never writes stdout', path: 'src/main.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false, cwd: 'read' },
  { name: 'the write CLI builder never writes stdout', path: 'src/composition/write-program.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
  { name: 'control: a write test may write stdout', path: 'src/composition/write-program.test.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: false, strict: false },
  { name: 'stdout is banned by default, not only in MCP files', path: 'src/composition/run-registry-command.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
  { name: 'control: the CLI builder owns stdout', path: 'src/composition/cli-shell.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: false, strict: false },
  { name: 'the CLI builder never imports the registry', path: 'src/composition/cli-shell.ts', source: REGISTRY_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the MCP gateway never imports a lifecycle command', path: 'src/composition/mcp-shell.ts', source: LOGIN_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'run-registry-command never imports the single package', path: 'src/composition/run-registry-command.ts', source: SINGLE_PACKAGE_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the bin placeholders never take the single package names', path: 'src/use-cases/commands/bin-names.ts', source: SINGLE_PACKAGE_NAMES_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the docs builders never import the lifecycle entries', path: 'src/use-cases/commands/docs-builders.ts', source: LIFECYCLE_ENTRIES_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the command framework never imports the registry', path: 'src/use-cases/commands/resolve-command.ts', source: SIBLING_REGISTRY_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'a core docs builder keeps the read-only ban', path: 'src/use-cases/commands/docs-builders.ts', source: WRITE_GRAPH_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the read graph keeps the core ban', path: 'src/infra/read-graph.ts', source: REGISTRY_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the CLI builder never takes the registry through the package entry', path: 'src/composition/cli-shell.ts', source: BARREL_REGISTRY_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the CLI builder never imports the registry folder', path: 'src/composition/cli-shell.ts', source: REGISTRY_DIRECTORY_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'a shared helper core owns never imports the registry', path: 'src/use-cases/commands/quote-boundary.ts', source: SIBLING_REGISTRY_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the command framework never imports its own folder', path: 'src/use-cases/commands/resolve-command.ts', source: SIBLING_DIRECTORY_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the MCP gateway never imports the published package', path: 'src/composition/mcp-shell.ts', source: SELF_PACKAGE_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the CLI builder never loads the registry with a dynamic import', path: 'src/composition/cli-shell.ts', source: DYNAMIC_REGISTRY_IMPORT, rule: 'no-restricted-syntax', fires: true, strict: false },
  { name: 'a core docs builder never loads the registry with a dynamic import', path: 'src/use-cases/commands/docs-builders.ts', source: DYNAMIC_REGISTRY_IMPORT.replace('../use-cases/commands/index.ts', './index.ts'), rule: 'no-restricted-syntax', fires: true, strict: false },
  { name: 'control: the core file system loads its Node half dynamically', path: 'src/infra/filesystem-bun.ts', source: DYNAMIC_SIBLING_IMPORT, rule: 'no-restricted-syntax', fires: false, strict: false },
  { name: 'the auth ladder never loads the browser', path: 'src/infra/auth.ts', source: BROWSER_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'control: the browser half of auth loads it', path: 'src/infra/auth-browser.ts', source: BROWSER_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'the token helper never loads winston', path: 'src/composition/token-helper.ts', source: WINSTON_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: "the token helper never loads auth's MCP server builder", path: 'src/composition/token-helper.ts', source: AUTH_SERVER_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: "the token helper never loads auth's CLI builder", path: 'src/composition/token-helper.ts', source: AUTH_PROGRAM_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the token helper never loads the MCP SDK', path: 'src/composition/token-helper.ts', source: MCP_TYPES_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the token helper never loads zod', path: 'src/composition/token-helper.ts', source: ZOD_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'control: an auth file outside the token helper loads winston', path: 'src/composition/auth-cli.ts', source: WINSTON_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'a read command never imports the write graph', path: 'src/use-cases/commands/list-drives.ts', source: WRITE_GRAPH_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'a read command never takes the write graph type', path: 'src/use-cases/commands/get-schedule.ts', source: WRITE_GRAPH_TYPE_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'a read helper never builds the client of the composition root', path: 'src/use-cases/commands/fetch-raw-bytes.ts', source: ROOT_CLIENT_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the read CLI never imports the write graph', path: 'src/composition/cli.ts', source: CORE_WRITE_GRAPH_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the read client never takes the write graph type', path: 'src/composition/read-client.ts', source: CORE_WRITE_GRAPH_TYPE_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the read MCP server never builds the client of the composition root', path: 'src/composition/mcp.ts', source: COMPOSITION_CLIENT_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'control: the composition root builds the client', path: 'src/composition/build-deps.ts', source: COMPOSITION_CLIENT_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'the read graph never imports the write graph', path: 'src/infra/read-graph.ts', source: SIBLING_WRITE_GRAPH_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'production code never imports a shared test fake through the core package', path: 'src/composition/single-package.ts', source: CORE_FAKE_IMPORT, rule: 'no-restricted-imports', fires: true, strict: false },
  { name: 'production code never imports a shared test fake by a relative path', path: 'src/composition/cli-shell.ts', source: SIBLING_FAKE_IMPORT, rule: 'no-restricted-imports', fires: true, strict: false },
  { name: 'control: a test imports the shared test fakes', path: 'src/composition/cli-shell.test.ts', source: SIBLING_FAKE_IMPORT, rule: 'no-restricted-imports', fires: false, strict: false },
  { name: 'strict lint rejects a needless assertion', path: 'src/composition/single-package.ts', source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: true, strict: true },
  { name: 'strict lint rejects a string rejection', path: 'src/composition/single-package.ts', source: STRING_REJECTION, rule: '@typescript-eslint/prefer-promise-reject-errors', fires: true, strict: true },
  { name: 'control: plain lint leaves type-aware rules off', path: 'src/composition/single-package.ts', source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: false, strict: false },
  // Each workspace folder spreads the root's eslint.base.js; its type-aware block
  // names `src/**`, which must resolve inside the folder.
  ...WORKSPACE_FILES.map(([cwd, path]): Fixture => ({ name: `${cwd}: strict lint rejects a needless assertion`, path, source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: true, strict: true, cwd })),
  { name: 'control: core: plain lint leaves type-aware rules off', path: 'src/domain/result.ts', source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: false, strict: false, cwd: 'core' },
  ...WORKSPACE_FILES.flatMap((file, i) => baseBanFixtures(file, WORKSPACE_TESTS[i] ?? file)),
  ...USE_CASE_FILES.flatMap(tryBanFixtures),
  ...USE_CASE_FILES.flatMap(layerZoneFixtures),
  ...CORE_LAYER_FIXTURES,
  { name: 'control: read: an infra adapter may catch', path: 'src/infra/mammoth-adapter.ts', source: TRY_CATCH, rule: 'no-restricted-syntax', fires: false, strict: false, cwd: 'read' },
  // Core's own no-restricted-syntax blocks replace the base's options, so they
  // repeat the call-assertion ban: one file of each block.
  { name: 'core: a core file keeps the call-assertion ban', path: 'src/composition/cli-shell.ts', source: CALL_ASSERTION_IN_SOURCE, rule: 'no-restricted-syntax', fires: true, strict: false, cwd: 'core' },
  { name: 'core: a read-only core file keeps the call-assertion ban', path: 'src/use-cases/commands/docs-builders.ts', source: CALL_ASSERTION_IN_SOURCE, rule: 'no-restricted-syntax', fires: true, strict: false, cwd: 'core' },
];

type LintMessage = { readonly ruleId: string | null; readonly message: string };
type LintResult = { readonly messages: ReadonlyArray<LintMessage> };
type Outcome = { readonly messages: ReadonlyArray<LintMessage> } | { readonly failure: string };

// LINT_STRICT is always removed: a strict fixture must get it from the
// lint:strict script, and a plain one must run without it.
const envWithoutStrict = (): Record<string, string | undefined> => {
  const env: Record<string, string | undefined> = { ...process.env };
  delete env['LINT_STRICT'];
  return env;
};

const command = (fixture: Fixture): ReadonlyArray<string> => {
  const stdinArgs = ['--stdin', '--stdin-filename', fixture.path, '--format', 'json'];
  return fixture.strict ? ['bun', 'run', 'lint:strict', ...stdinArgs] : ['bunx', 'eslint', ...stdinArgs];
};

const lint = async (fixture: Fixture): Promise<Outcome> => {
  const proc = Bun.spawn([...command(fixture)], {
    stdin: new TextEncoder().encode(fixture.source),
    cwd: fixture.cwd,
    stdout: 'pipe',
    stderr: 'pipe',
    env: envWithoutStrict(),
  });
  const [stdout, stderr] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text()]);
  const status = await proc.exited;
  try {
    const results = JSON.parse(stdout) as ReadonlyArray<LintResult>;
    return { messages: results.flatMap((r) => r.messages) };
  } catch {
    return { failure: `eslint exited ${status} without a JSON report: ${stderr.trim().slice(0, 300)}` };
  }
};

const ruleOf = (m: LintMessage): string | null => (m.ruleId === null && m.message.includes("'noInlineConfig' setting") ? INLINE_CONFIG : m.ruleId);

// The folders that a guarded file of the root src/ moves to (package split, phases 3 to 6).
const MOVED_TO: ReadonlyArray<string> = ['core', 'auth', 'write', 'read'];

// A guarded file that moved is linted from its folder, with that folder's
// config: the rule must hold there too. A path in two folders (src/index.ts,
// src/main.ts) names no one file, so the fixture must say its folder.
const located = async (fixture: Fixture): Promise<Fixture | { readonly failure: string }> => {
  if (fixture.cwd !== undefined || (await Bun.file(fixture.path).exists())) return fixture;
  const found = await Promise.all(MOVED_TO.map(async (cwd) => ((await Bun.file(`${cwd}/${fixture.path}`).exists()) ? [cwd] : [])));
  const folders = found.flat();
  if (folders.length > 1) return { failure: `${fixture.path} exists in ${folders.join(', ')}: give the fixture the cwd of its folder` };
  return folders[0] === undefined ? fixture : { ...fixture, cwd: folders[0] };
};

const check = async (listed: Fixture): Promise<boolean> => {
  const placed = await located(listed);
  if ('failure' in placed) {
    console.log(`  FAIL  ${listed.name}  (${listed.path}, ${listed.rule})\n        ${placed.failure}`);
    return false;
  }
  const fixture = placed;
  const label = `${fixture.name}  (${fixture.cwd === undefined ? '' : `${fixture.cwd}/`}${fixture.path}, ${fixture.rule})`;
  if (!(await Bun.file(fixture.cwd === undefined ? fixture.path : `${fixture.cwd}/${fixture.path}`).exists())) {
    console.log(`  FAIL  ${label}\n        ${fixture.path} no longer exists: move this fixture and the rule's glob in eslint.config.js with it`);
    return false;
  }
  const outcome = await lint(fixture);
  if ('failure' in outcome) {
    console.log(`  FAIL  ${label}\n        ${outcome.failure}`);
    return false;
  }
  // A message without a rule is a parse error or an "ignored file" warning: the
  // text was never linted, so neither a firing nor a silent rule proves anything.
  // The noInlineConfig report has no rule either, and has its own label.
  const unlinted = outcome.messages.find((m) => ruleOf(m) === null);
  if (unlinted) {
    console.log(`  FAIL  ${label}\n        not linted: ${unlinted.message}`);
    return false;
  }
  const fired = outcome.messages.some((m) => ruleOf(m) === fixture.rule);
  const held = fired === fixture.fires;
  console.log(`  ${held ? 'ok  ' : 'FAIL'}  ${label} ${fired ? 'fired' : 'silent'}`);
  return held;
};

// At most this many eslint processes run at once. Each one takes about 180 MB,
// and 86 fixtures at once filled the 16 GB of a 4-core machine: the run took
// more than 30 minutes instead of about one.
const PARALLEL = Math.max(2, navigator.hardwareConcurrency);

const checkAll = async (fixtures: ReadonlyArray<Fixture>): Promise<ReadonlyArray<boolean>> => {
  const held: boolean[] = [];
  let next = 0;
  const worker = async (): Promise<void> => {
    for (let i = next++; i < fixtures.length; i = next++) {
      const fixture = fixtures[i];
      if (fixture !== undefined) held[i] = await check(fixture);
    }
  };
  await Promise.all(Array.from({ length: PARALLEL }, worker));
  return held;
};

const main = async (): Promise<number> => {
  console.log('lint fixtures:');
  const outcomes = await checkAll(FIXTURES);
  const failed = outcomes.filter((held) => !held).length;
  if (failed > 0) {
    console.error(`lint fixtures: ${failed} fixture(s) failed. A path-scoped rule stopped matching its files, or fires where it must not.`);
    return 1;
  }
  console.log('lint fixtures: every path-scoped rule fires on its fixture and only there.');
  return 0;
};

process.exit(await main());
