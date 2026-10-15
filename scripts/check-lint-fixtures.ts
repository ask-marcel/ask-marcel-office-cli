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
const WRITE_GRAPH_IMPORT = "import { createWriteGraph } from '../../infra/write-graph.ts';\nexport const f = (): unknown => createWriteGraph;\n";
const WRITE_GRAPH_TYPE_IMPORT = "import type { WriteGraph } from '../../infra/write-graph.ts';\nexport type F = WriteGraph;\n";
const FULL_CLIENT_IMPORT = "import { createGraphClient } from '../../infra/graph-client.ts';\nexport const f = (): unknown => createGraphClient;\n";
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

// One real source file per workspace folder, relative to the folder.
const WORKSPACE_FILES: ReadonlyArray<readonly [string, string]> = [
  ['core', 'src/domain/result.ts'],
  ['auth', 'src/composition/bin.ts'],
  ['read', 'src/composition/bin.ts'],
  ['write', 'src/composition/bin.ts'],
];

const FIXTURES: ReadonlyArray<Fixture> = [
  { name: 'the MCP server never writes stdout', path: 'src/composition/mcp.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
  { name: 'the MCP gateway never writes stdout', path: 'src/composition/mcp-shell.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
  { name: 'the auth MCP tools never write stdout', path: 'src/composition/auth-mcp.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
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
  { name: 'control: the composition root wires the registry', path: 'src/composition/single-package.ts', source: REGISTRY_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'control: the auth CLI commands import login', path: 'src/composition/auth-cli.ts', source: LOGIN_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'the auth ladder never loads the browser', path: 'src/infra/auth.ts', source: BROWSER_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'control: the browser half of auth loads it', path: 'src/infra/auth-browser.ts', source: BROWSER_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'the token helper never loads winston', path: 'src/composition/token-helper.ts', source: WINSTON_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'control: the logger adapter wraps winston', path: 'src/infra/logger.ts', source: WINSTON_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'a read command never imports the write graph', path: 'src/use-cases/commands/list-drives.ts', source: WRITE_GRAPH_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'a read command never takes the write graph type', path: 'src/use-cases/commands/get-schedule.ts', source: WRITE_GRAPH_TYPE_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'a read helper never builds the full client', path: 'src/use-cases/commands/fetch-raw-bytes.ts', source: FULL_CLIENT_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'the read graph never imports the write graph', path: 'src/infra/read-graph.ts', source: SIBLING_WRITE_GRAPH_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'control: a write command takes the write graph', path: 'src/use-cases/commands/create-mail-draft.ts', source: WRITE_GRAPH_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'strict lint rejects a needless assertion', path: 'src/domain/result.ts', source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: true, strict: true },
  { name: 'strict lint rejects a string rejection', path: 'src/domain/result.ts', source: STRING_REJECTION, rule: '@typescript-eslint/prefer-promise-reject-errors', fires: true, strict: true },
  { name: 'control: plain lint leaves type-aware rules off', path: 'src/domain/result.ts', source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: false, strict: false },
  // Each workspace folder spreads the root's eslint.base.js; its type-aware block
  // names `src/**`, which must resolve inside the folder.
  ...WORKSPACE_FILES.map(([cwd, path]): Fixture => ({ name: `${cwd}: strict lint rejects a needless assertion`, path, source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: true, strict: true, cwd })),
  { name: 'control: core: plain lint leaves type-aware rules off', path: 'src/domain/result.ts', source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: false, strict: false, cwd: 'core' },
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

const check = async (fixture: Fixture): Promise<boolean> => {
  const label = `${fixture.name}  (${fixture.path}, ${fixture.rule})`;
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
  const unlinted = outcome.messages.find((m) => m.ruleId === null);
  if (unlinted) {
    console.log(`  FAIL  ${label}\n        not linted: ${unlinted.message}`);
    return false;
  }
  const fired = outcome.messages.some((m) => m.ruleId === fixture.rule);
  const held = fired === fixture.fires;
  console.log(`  ${held ? 'ok  ' : 'FAIL'}  ${label} ${fired ? 'fired' : 'silent'}`);
  return held;
};

const main = async (): Promise<number> => {
  console.log('lint fixtures:');
  const outcomes = await Promise.all(FIXTURES.map(check));
  const failed = outcomes.filter((held) => !held).length;
  if (failed > 0) {
    console.error(`lint fixtures: ${failed} fixture(s) failed. A path-scoped rule stopped matching its files, or fires where it must not.`);
    return 1;
  }
  console.log('lint fixtures: every path-scoped rule fires on its fixture and only there.');
  return 0;
};

process.exit(await main());
