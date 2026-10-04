#!/usr/bin/env bun
/*
 * Packed-install smoke: the tarball npm would serve installs and runs outside
 * the repo.
 *
 * Every other gate runs inside the repo, where node_modules holds every
 * dependency the repo ever declared. A package whose own `dependencies` miss a
 * module its bundle leaves external passes all of them and then fails on a
 * user's machine. This packs the package, installs the tarball into an empty
 * directory under the OS temp dir (checked to have no node_modules above it),
 * and runs the installed copy:
 *
 *   1. every package the bundle imports at run time (its external static and
 *      dynamic imports, read from the installed dist) loads under node and bun,
 *      including the ones no converter reaches, such as login's playwright; and
 *      a repo devDependency the tarball never installs fails to load, which
 *      re-proves on every run that the install is isolated (rule 15.10);
 *   2. `--version` answers the manifest's version, through the bin path under
 *      node and bun and through the installed `.bin` shim;
 *   3. the library loads by package name, and its `commands.json` subpath
 *      resolves, under node and bun;
 *   4. qa-bundle-smoke.ts, pointed at the installed copy, passes: every
 *      converter and the MCP stdout-purity handshake, under node and bun.
 *
 * Run after the build: `bun run build && bun scripts/qa-packed-install-smoke.ts`.
 * Needs the registry, since the tarball's dependencies are installed fresh.
 * Exit non-zero on any failure.
 */
import { existsSync, mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join, sep } from 'node:path';
import { spawnSync } from 'node:child_process';

type Manifest = {
  readonly name: string;
  readonly version: string;
  readonly bin?: Record<string, string>;
};

const RUNTIMES = ['node', 'bun'] as const;
const PROBE_TIMEOUT_MS = 60000;
const INSTALL_TIMEOUT_MS = 300000;

const fail = (message: string): never => {
  throw new Error(message);
};

const run = (cmd: string, args: ReadonlyArray<string>, cwd: string, timeout = PROBE_TIMEOUT_MS, env: NodeJS.ProcessEnv = process.env): string => {
  const p = spawnSync(cmd, args, { cwd, timeout, env, maxBuffer: 64 * 1024 * 1024 });
  if (p.error !== undefined || p.signal !== null || p.status !== 0) {
    const cause = p.error?.message ?? (p.signal === null ? `exit ${p.status}` : `signal ${p.signal}`);
    fail(`\`${cmd} ${args.join(' ').slice(0, 120)}\` failed (${cause}): ${(p.stderr?.toString() || p.stdout?.toString() || '').slice(0, 400)}`);
  }
  return p.stdout?.toString().trim() ?? '';
};

// Evaluates an ES module snippet with the given runtime, resolving from cwd.
const evaluate = (rt: (typeof RUNTIMES)[number], code: string, cwd: string): string =>
  run(rt, rt === 'node' ? ['--input-type=module', '-e', code] : ['-e', code], cwd);

// The packages the bundle imports at run time: whole import statements and
// dynamic imports only (the bundle also carries generated code with
// `require("...")` inside strings, which is not an import).
const bundleExternals = async (root: string): Promise<ReadonlyArray<string>> => {
  const texts = await Promise.all(['dist/cli.js', 'dist/index.js', 'dist/token.js'].map((f) => Bun.file(join(root, f)).text()));
  const found = texts.flatMap((t) => [...t.matchAll(/^import [^;]* from "([^"]+)";$|^import "([^"]+)";$|import\("([^"]+)"\)/gm)].map((m) => m[1] ?? m[2] ?? m[3] ?? ''));
  return [...new Set(found)].filter((s) => s !== '' && !s.startsWith('.') && !s.startsWith('node:') && !builtinModules.includes(s)).toSorted();
};

// "Outside the repo" must hold for the result to mean anything: a node_modules
// in any ancestor of the consumer would supply a dependency the tarball forgot.
const assertIsolated = (consumer: string, repo: string): void => {
  if (consumer.startsWith(repo + sep)) fail(`the consumer ${consumer} sits inside the repo`);
  let dir = dirname(consumer);
  for (;;) {
    if (existsSync(join(dir, 'node_modules'))) fail(`${join(dir, 'node_modules')} could supply dependencies the tarball forgot; set TMPDIR elsewhere`);
    const up = dirname(dir);
    if (up === dir) return;
    dir = up;
  }
};

const repo = process.cwd();
const work = mkdtempSync(join(tmpdir(), 'packed-install-'));
let exitCode = 0;
try {
  const tarball = run('bun', ['pm', 'pack', '--destination', work, '--quiet'], repo);
  console.log(`packed-install smoke: packed ${tarball}`);

  const consumer = join(work, 'consumer');
  mkdirSync(consumer);
  assertIsolated(consumer, repo);
  await Bun.write(join(consumer, 'package.json'), JSON.stringify({ name: 'packed-install-consumer', private: true }));
  run('bun', ['add', tarball], consumer, INSTALL_TIMEOUT_MS);

  const source = (await Bun.file(join(repo, 'package.json')).json()) as Manifest;
  const root = join(consumer, 'node_modules', source.name);
  const installed = (await Bun.file(join(root, 'package.json')).json()) as Manifest;

  const externals = await bundleExternals(root);
  if (externals.length === 0) fail('found no external imports in the installed bundle; has the build output changed shape?');
  for (const rt of RUNTIMES) evaluate(rt, `for (const m of ${JSON.stringify(externals)}) await import(m);`, root);
  console.log(`  ✓ ${externals.length} external imports load from the installed copy under node and bun: ${externals.join(', ')}`);

  // The red path, re-run every time: a package the repo has but the tarball
  // never installs must NOT load from the installed copy. If it loads, something
  // outside the consumer supplies modules, and every check here proves nothing.
  const repoDevDependencies = Object.keys(((await Bun.file(join(repo, 'package.json')).json()) as { devDependencies?: Record<string, string> }).devDependencies ?? {});
  const undeclared = repoDevDependencies.find((d) => !existsSync(join(consumer, 'node_modules', d))) ?? fail('no repo devDependency is absent from the install, so isolation cannot be checked');
  for (const rt of RUNTIMES) {
    const loaded = ((): boolean => {
      try {
        evaluate(rt, `await import(${JSON.stringify(undeclared)});`, root);
        return true;
      } catch {
        return false;
      }
    })();
    if (loaded) fail(`${undeclared}, which the tarball does not install, loads from the installed copy under ${rt}: the consumer is not isolated`);
  }
  console.log(`  ✓ ${undeclared}, which the tarball does not install, fails to load under node and bun (the consumer is isolated)`);

  // Every CLI run records where its token helper lives under HOME, so the bin
  // runs on a scratch home and never repoints the user's.
  const scratchHome = join(work, 'home');
  const binEnv = { ...process.env, HOME: scratchHome, USERPROFILE: scratchHome };
  const [binName, binPath] = Object.entries(installed.bin ?? {})[0] ?? fail('the packed manifest declares no bin');
  for (const rt of RUNTIMES) {
    const version = run(rt, [join(root, binPath), '--version'], consumer, PROBE_TIMEOUT_MS, binEnv);
    if (version !== installed.version) fail(`--version under ${rt} answered ${JSON.stringify(version)}, expected ${installed.version}`);
  }
  const shimVersion = run(join(consumer, 'node_modules', '.bin', binName), ['--version'], consumer, PROBE_TIMEOUT_MS, binEnv);
  if (shimVersion !== installed.version) fail(`the ${binName} shim answered ${JSON.stringify(shimVersion)}, expected ${installed.version}`);
  console.log(`  ✓ --version ${installed.version} under node, bun and the ${binName} shim`);

  const library = `const m = await import(${JSON.stringify(source.name)}); process.stdout.write(JSON.stringify({ commands: Object.keys(m.commands).length, manifest: import.meta.resolve(${JSON.stringify(`${source.name}/commands.json`)}) }));`;
  for (const rt of RUNTIMES) {
    const answer = JSON.parse(evaluate(rt, library, consumer)) as { commands: number; manifest: string };
    if (answer.commands === 0) fail(`the library loaded by name under ${rt} exports no commands`);
    if (!existsSync(new URL(answer.manifest))) fail(`${source.name}/commands.json resolved under ${rt} to a missing file: ${answer.manifest}`);
  }
  console.log(`  ✓ the library loads by name and ${source.name}/commands.json resolves, under node and bun`);

  const smoke = spawnSync('bun', [join(repo, 'scripts/qa-bundle-smoke.ts'), '--package-root', root], { cwd: repo, stdio: 'inherit', timeout: INSTALL_TIMEOUT_MS });
  if (smoke.error !== undefined || smoke.signal !== null || smoke.status !== 0) fail('the bundle smoke failed against the installed copy');
  console.log('packed-install smoke: the tarball installs and runs outside the repo ✓');
} catch (error) {
  console.error(`packed-install smoke: ${error instanceof Error ? error.message : String(error)}`);
  exitCode = 1;
} finally {
  rmSync(work, { recursive: true, force: true });
}
process.exit(exitCode);
