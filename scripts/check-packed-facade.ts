#!/usr/bin/env bun
/*
 * Packed-output check of the facade (cli/, ADR 0005). Run it with cli/ as the
 * working directory, after `bun run build` in auth/, read/, write/ and cli/:
 *
 *   bun run ../scripts/check-packed-facade.ts
 *
 * The facade starts the bin of an installed package, so nothing in the
 * workspace can show that an install of it works: the hoisted workspace
 * resolves the three packages for every check that runs inside it. This checks
 * what a person installs:
 *
 *   1. the tarball of `bun pm pack` holds package.json, LICENSE, README.md,
 *      CHANGELOG.md and the three bins, and no other file;
 *   2. its manifest names the three packages, each at the exact version of its
 *      folder, and no other dependency; its bins are the bins of the three
 *      packages;
 *   3. a clean consumer installs the local bundle (scripts/facade-bundle.ts:
 *      no registry has the three packages yet). Each bin answers --version
 *      with the version of its package and help-json with the name of its
 *      package (the three versions can be equal, the names cannot), under
 *      node and bun, with a temporary home and no ASKMARCEL_* variable.
 *
 * Every run also proves that step 3 can fail (rule 15.10): it removes one
 * package from the installed facade, and the bin of that package must then
 * fail and say that the package is not installed.
 *
 *   --selftest  proves that steps 1 and 2 reject a stray file, a missing bin
 *               file, a `workspace:*` version, a fourth dependency and a bin
 *               that the facade does not install, and that the bundle manifest
 *               holds the registry dependencies of the three packages and
 *               refuses two ranges for one dependency, without packing
 *               anything.
 */

import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { Manifest } from './facade-bundle.ts';
import { FACADE, PACKAGES, buildBundle, bundleManifest, packFolder, readManifest } from './facade-bundle.ts';

const FILES: ReadonlyArray<string> = ['package.json', 'LICENSE', 'README.md', 'CHANGELOG.md', 'dist/auth.js', 'dist/read.js', 'dist/write.js'];

const fileProblems = (files: ReadonlyArray<string>): ReadonlyArray<string> => [
  ...FILES.filter((file) => !files.includes(file)).map((file) => `the tarball misses ${file}`),
  ...files.filter((file) => !FILES.includes(file)).map((file) => `the tarball holds ${file}, which is not a file of the facade`),
];

const binNames = (manifest: Manifest): ReadonlyArray<string> => Object.keys(manifest.bin ?? {}).sort();

const manifestProblems = (facade: Manifest, packages: ReadonlyArray<Manifest>): ReadonlyArray<string> => {
  const dependencies = facade.dependencies ?? {};
  const names = packages.map((manifest) => manifest.name);
  const expected = packages.flatMap(binNames).sort();
  return [
    ...packages
      .filter((manifest) => dependencies[manifest.name] !== manifest.version)
      .map((manifest) => `dependencies must name ${manifest.name} at ${manifest.version}, not at ${dependencies[manifest.name] ?? 'no version'}`),
    ...Object.keys(dependencies)
      .filter((name) => !names.includes(name))
      .map((name) => `dependencies names ${name}, which is not one of the three packages`),
    ...(binNames(facade).join(' ') === expected.join(' ') ? [] : [`the bins must be ${expected.join(', ')}, not ${binNames(facade).join(', ') || 'none'}`]),
  ];
};

type Ran = { readonly ok: boolean; readonly stdout: string; readonly out: string };

// `stdout` is what a command answers. `out` adds stderr, for the text of a
// problem only: a warning on stderr is not an answer.
const run = (cmd: ReadonlyArray<string>, cwd: string, env?: Readonly<Record<string, string>>): Ran => {
  const proc = Bun.spawnSync([...cmd], { cwd, stdout: 'pipe', stderr: 'pipe', ...(env === undefined ? {} : { env: { ...env } }) });
  const stdout = proc.stdout.toString();
  return { ok: proc.exitCode === 0, stdout, out: `${stdout}${proc.stderr.toString()}` };
};

// The `package` field of a help-json answer, or undefined for an answer that is not JSON.
const packageOf = (answer: string): string | undefined => {
  try {
    return (JSON.parse(answer) as { readonly package?: string }).package;
  } catch {
    return undefined;
  }
};

// The caller's environment with a temporary home and no ASKMARCEL_* variable:
// the auth bin records the token helper locator, and it must not write the
// real ~/.ask-marcel/ (as in scripts/check-packed-package.ts).
const isolatedEnv = (home: string): Readonly<Record<string, string>> => ({
  ...Object.fromEntries(Object.entries(process.env).flatMap(([key, value]) => (value === undefined || key.startsWith('ASKMARCEL_') ? [] : [[key, value] as const]))),
  HOME: home,
  USERPROFILE: home,
});

// The bin file of the facade that has the bin name of a package.
const facadeBin = (installed: string, facade: Manifest, manifest: Manifest): string => join(installed, facade.bin?.[binNames(manifest)[0] ?? ''] ?? '');

const runtimeProblems = (consumer: string, installed: string, facade: Manifest, packages: ReadonlyArray<Manifest>): ReadonlyArray<string> => {
  const env = isolatedEnv(mkdtempSync(join(consumer, 'home-')));
  const versions = packages.flatMap((manifest) =>
    ['node', 'bun'].flatMap((runtime) => {
      const bin = facadeBin(installed, facade, manifest);
      const version = run([runtime, bin, '--version'], consumer, env);
      const help = run([runtime, bin, 'help-json'], consumer, env);
      return [
        ...(version.ok && version.stdout.trim() === manifest.version
          ? []
          : [`${runtime}: the bin of ${manifest.name} did not answer --version with ${manifest.version}: ${version.out.slice(0, 300)}`]),
        ...(help.ok && packageOf(help.stdout) === manifest.name ? [] : [`${runtime}: the bin of ${manifest.name} did not answer help-json as that package: ${help.out.slice(0, 300)}`]),
      ];
    })
  );
  // The control: with one package removed, its bin must fail and name the cause.
  const removed = packages.at(-1);
  if (removed === undefined) return [...versions, 'no package to remove for the control'];
  rmSync(join(installed, 'node_modules', removed.name), { recursive: true, force: true });
  const control = run(['node', facadeBin(installed, facade, removed), '--version'], consumer, env);
  const failed = !control.ok && control.out.includes(`the package ${removed.name} is not installed`);
  return [...versions, ...(failed ? [] : [`the control passed: with ${removed.name} removed, its bin must fail and say so: ${control.out.slice(0, 300)}`])];
};

const consumerProblems = (bundle: string, facade: Manifest, packages: ReadonlyArray<Manifest>): ReadonlyArray<string> => {
  const consumer = mkdtempSync(join(tmpdir(), 'facade-consumer-'));
  try {
    writeFileSync(join(consumer, 'package.json'), JSON.stringify({ name: 'facade-consumer', private: true, type: 'module' }));
    const install = run(['bun', 'add', bundle], consumer);
    if (!install.ok) return [`the bundle does not install: ${install.out.slice(0, 400)}`];
    return runtimeProblems(consumer, join(consumer, 'node_modules', facade.name), facade, packages);
  } finally {
    rmSync(consumer, { recursive: true, force: true });
  }
};

const expectProblem = (label: string, problems: ReadonlyArray<string>, part: string): ReadonlyArray<string> =>
  problems.some((problem) => problem.includes(part)) ? [] : [`SELFTEST FAILED, ${label}: no problem names "${part}" (got: ${problems.join('; ') || 'none'})`];

const throws = (action: () => unknown): boolean => {
  try {
    action();
    return false;
  } catch {
    return true;
  }
};

const selftest = (): number => {
  const one: Manifest = { name: '@x/one', version: '1.2.0', bin: { 'x-one': './dist/cli.js' }, dependencies: { zod: '^4.0.0', left: '^1.0.0' } };
  const two: Manifest = { name: '@x/two', version: '2.0.0', bin: { 'x-two': './dist/cli.js' }, dependencies: { zod: '^4.0.0' } };
  const facade: Manifest = {
    name: 'x',
    version: '3.0.0',
    bin: { 'x-one': './dist/one.js', 'x-two': './dist/two.js' },
    dependencies: { '@x/one': '1.2.0', '@x/two': '2.0.0' },
    devDependencies: { core: '1.0.0' },
  };
  const bundle = bundleManifest(facade, [one, two]);
  const failures = [
    ...(fileProblems(FILES).length === 0 ? [] : ['SELFTEST FAILED: the right file list has a problem']),
    ...expectProblem('a stray file', fileProblems([...FILES, 'dist/index.js']), 'holds dist/index.js'),
    ...expectProblem('a missing bin file', fileProblems(FILES.filter((file) => file !== 'dist/read.js')), 'misses dist/read.js'),
    ...(manifestProblems(facade, [one, two]).length === 0 ? [] : ['SELFTEST FAILED: the right manifest has a problem']),
    ...expectProblem('a workspace version', manifestProblems({ ...facade, dependencies: { '@x/one': 'workspace:*', '@x/two': '2.0.0' } }, [one, two]), '@x/one at 1.2.0'),
    ...expectProblem('a fourth dependency', manifestProblems({ ...facade, dependencies: { ...facade.dependencies, zod: '^4.0.0' } }, [one, two]), 'names zod'),
    ...expectProblem('a bin that the facade does not install', manifestProblems({ ...facade, bin: { 'x-one': './dist/one.js' } }, [one, two]), 'the bins must be x-one, x-two'),
    ...(JSON.stringify(bundle.dependencies) === JSON.stringify({ zod: '^4.0.0', left: '^1.0.0' }) ? [] : ['SELFTEST FAILED: the bundle does not hold the registry dependencies']),
    ...(JSON.stringify(bundle['bundledDependencies']) === JSON.stringify(['@x/one', '@x/two']) ? [] : ['SELFTEST FAILED: the bundle does not name the packages inside it']),
    ...(bundle.devDependencies === undefined ? [] : ['SELFTEST FAILED: the bundle keeps devDependencies']),
    ...(throws(() => bundleManifest(facade, [one, { ...two, dependencies: { zod: '^3.0.0' } }])) ? [] : ['SELFTEST FAILED: two ranges for one dependency are accepted']),
  ];
  for (const failure of failures) console.error(`check-packed-facade: ${failure}`);
  if (failures.length === 0) console.log('check-packed-facade: selftest passed (11 cases).');
  return failures.length === 0 ? 0 : 1;
};

const main = async (): Promise<number> => {
  const dir = mkdtempSync(join(tmpdir(), 'facade-packed-'));
  try {
    const tarball = packFolder(FACADE, dir);
    const files = run(['tar', '-tzf', tarball], dir)
      .stdout.split('\n')
      .map((path) => path.trim().replace(/^package\//, ''))
      .filter((path) => path !== '' && !path.endsWith('/'));
    run(['tar', '-xzf', tarball], dir);
    const facade = (await Bun.file(join(dir, 'package', 'package.json')).json()) as Manifest;
    const packages = await Promise.all(PACKAGES.map(readManifest));
    const bundle = await buildBundle(join(dir, 'bundle'));
    const problems = [...fileProblems(files), ...manifestProblems(facade, packages), ...consumerProblems(bundle, facade, packages)];
    for (const problem of problems) console.error(`check-packed-facade: ${problem}`);
    if (problems.length > 0) return 1;
    console.log(
      `check-packed-facade: ${files.length} files, ${packages.length} packages at their exact versions, and ${packages.length} bins that start under node and bun from a clean install of the bundle; the control failed as it must.`
    );
    return 0;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
};

process.exit(process.argv.includes('--selftest') ? selftest() : await main());
