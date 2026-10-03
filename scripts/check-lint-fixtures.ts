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
};

const STDOUT_WRITE = "export const f = (): void => {\n  process.stdout.write('banner');\n};\n";
const NEEDLESS_ASSERTION = 'const n = 1;\nexport const f = (): number => n as number;\n';
const STRING_REJECTION = "export const f = (): Promise<never> => Promise.reject('nope');\n";
const BROWSER_IMPORT = "import { createBrowserAuth } from './browser-auth.ts';\nexport const f = (): unknown => createBrowserAuth;\n";

const FIXTURES: ReadonlyArray<Fixture> = [
  { name: 'the MCP server never writes stdout', path: 'src/composition/mcp.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: true, strict: false },
  { name: 'control: the CLI owns stdout', path: 'src/composition/cli.ts', source: STDOUT_WRITE, rule: 'no-restricted-properties', fires: false, strict: false },
  { name: 'the auth ladder never loads the browser', path: 'src/infra/auth.ts', source: BROWSER_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: true, strict: false },
  { name: 'control: the browser half of auth loads it', path: 'src/infra/auth-browser.ts', source: BROWSER_IMPORT, rule: '@typescript-eslint/no-restricted-imports', fires: false, strict: false },
  { name: 'strict lint rejects a needless assertion', path: 'src/domain/result.ts', source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: true, strict: true },
  { name: 'strict lint rejects a string rejection', path: 'src/domain/result.ts', source: STRING_REJECTION, rule: '@typescript-eslint/prefer-promise-reject-errors', fires: true, strict: true },
  { name: 'control: plain lint leaves type-aware rules off', path: 'src/domain/result.ts', source: NEEDLESS_ASSERTION, rule: '@typescript-eslint/no-unnecessary-type-assertion', fires: false, strict: false },
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
  if (!(await Bun.file(fixture.path).exists())) {
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
