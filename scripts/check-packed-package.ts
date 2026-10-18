#!/usr/bin/env bun
/*
 * Packed-output check of one workspace package (auth/, read/, write/). Run it
 * with the folder as the working directory, after `bun run build`:
 *
 *   bun run ../scripts/check-packed-package.ts
 *
 * Core is private and never published, and each bundle inlines it. Nothing in
 * the repo can see a tarball that still points at core: the hoisted workspace
 * resolves `@ask-marcel/office-core` for every check that runs inside it. So
 * this packs the package with `bun pm pack` and checks the tarball itself:
 *
 *   1. the files: package.json, README.md, LICENSE, CHANGELOG.md, each file
 *      that `bin` and `exports` name, each bundle that the build makes
 *      (scripts/package-targets.ts: dist/token.js in auth), and declarations
 *      under dist/; no other;
 *   2. the manifest: `@ask-marcel/office-core` only in devDependencies (bun
 *      publish rewrites `workspace:*` to core's version, and an install would
 *      then 404 on the never-published core);
 *   3. the declarations: none mentions `office-core` or `../core`;
 *   4. a clean consumer (an empty folder under the OS temp dir, no node_modules
 *      above it) installs the tarball. The bin answers --version and the library
 *      loads by name, under node and bun, with a temporary HOME and
 *      USERPROFILE and no ASKMARCEL_* variable. Then `tsc --noEmit --skipLibCheck
 *      false` (module nodenext) runs on a file that uses each exported type once
 *      as it should and once wrongly, and calls each exported value once wrongly:
 *      it must report exactly the expected error on each wrong line, so a missing
 *      declaration (TS2307) or a type or value that became `any` (no error at
 *      all) fails the check.
 *
 * Every run also proves that step 4 can fail (rule 15.10): it replaces the
 * installed index.d.ts with one that imports core (must fail with TS2307), and
 * with one whose types and values are `any` (must fail as a silent pass).
 *
 *   --selftest  proves steps 1 to 3 reject a stray file, core in dependencies,
 *               and a declaration that names core, and that the consumer file
 *               expects an error on each wrong line and refuses an export with no
 *               wrong use, without packing anything.
 *
 * An exported type with no entry in MISUSES, or an exported value with no entry
 * in VALUE_MISUSES, fails the check: a new public export comes with its own
 * wrong use.
 */

import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import ts from 'typescript';
import { bundleTargets } from './package-targets.ts';

type Manifest = {
  readonly name: string;
  readonly version: string;
  readonly bin?: Readonly<Record<string, string>>;
  readonly exports?: Readonly<Record<string, string | Readonly<Record<string, string>>>>;
  readonly [field: string]: unknown;
};

type Misuse = { readonly use: string; readonly fits: string; readonly misuse: string; readonly code: number };
type ValueMisuse = { readonly misuse: string; readonly code: number };

// One right and one wrong use per exported type name, the same in every package.
const MISUSES: Readonly<Record<string, Misuse>> = {
  Result: {
    use: 'Result<number, string>',
    fits: '{ ok: true, value: 1 }',
    misuse: "{ ok: true, value: 'one' }",
    code: 2322,
  },
  AuthPaths: {
    use: 'AuthPaths',
    fits: "{ tokenCache: '/h/.ask-marcel/token-cache.json', browserProfile: '/h/.ask-marcel/browser-profile' }",
    misuse: "{ tokenCache: '/h/.ask-marcel/token-cache.json' }",
    code: 2741,
  },
};

// One wrong call per exported value name. The consumer imports every value of
// the package, so a row may use another exported value.
const VALUE_MISUSES: Readonly<Record<string, ValueMisuse>> = {
  resolveAuthPaths: { misuse: 'resolveAuthPaths(1, {})', code: 2345 },
};

const CORE = '@ask-marcel/office-core';
const RUNTIME_FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies', 'bundleDependencies', 'bundledDependencies'];
const TOP_FILES = ['package.json', 'README.md', 'LICENSE', 'CHANGELOG.md'];
const TSC = resolve(import.meta.dir, '../node_modules/typescript/bin/tsc');

const withoutDot = (path: string): string => (path.startsWith('./') ? path.slice(2) : path);

const namedFiles = (manifest: Manifest): ReadonlyArray<string> => {
  const fromExports = Object.values(manifest.exports ?? {}).flatMap((target) => (typeof target === 'string' ? [target] : Object.values(target)));
  return [...Object.values(manifest.bin ?? {}), ...fromExports].map(withoutDot);
};

const fileProblems = (files: ReadonlyArray<string>, manifest: Manifest, bundles: ReadonlyArray<string>): ReadonlyArray<string> => {
  const required = [...new Set([...TOP_FILES, ...namedFiles(manifest), ...bundles])];
  const missing = required.filter((path) => !files.includes(path)).map((path) => `missing from the tarball: ${path}`);
  const stray = files.filter((path) => !required.includes(path) && !/^dist\/.+\.d\.ts$/.test(path)).map((path) => `not expected in the tarball: ${path}`);
  return [...missing, ...stray];
};

const manifestProblems = (manifest: Manifest): ReadonlyArray<string> =>
  RUNTIME_FIELDS.filter((field) => JSON.stringify(manifest[field] ?? {}).includes(CORE)).map((field) => `${CORE} is in ${field}: only devDependencies may name it`);

const declarationProblems = (declarations: Readonly<Record<string, string>>): ReadonlyArray<string> =>
  Object.entries(declarations)
    .filter(([, text]) => text.includes('office-core') || text.includes('../core'))
    .map(([path]) => `${path} names the private core (office-core or ../core)`);

const run = (cmd: ReadonlyArray<string>, cwd: string, env?: Readonly<Record<string, string>>): { readonly ok: boolean; readonly out: string } => {
  const proc = Bun.spawnSync([...cmd], { cwd, stdout: 'pipe', stderr: 'pipe', ...(env === undefined ? {} : { env: { ...env } }) });
  return { ok: proc.exitCode === 0, out: `${proc.stdout.toString()}${proc.stderr.toString()}` };
};

// The environment of the packed bin and library: the caller's, with a temporary
// home and no ASKMARCEL_* variable. A bin that records the token helper locator
// must not write the real ~/.ask-marcel/, and a token variable of the person who
// runs the check must not change what the bin answers.
const isolatedEnv = (home: string): Readonly<Record<string, string>> => ({
  ...Object.fromEntries(Object.entries(process.env).flatMap(([key, value]) => (value === undefined || key.startsWith('ASKMARCEL_') ? [] : [[key, value] as const]))),
  HOME: home,
  USERPROFILE: home,
});

const hasNodeModulesAbove = (dir: string): boolean => {
  const parent = dirname(dir);
  if (parent === dir) return false;
  return existsSync(join(parent, 'node_modules')) || hasNodeModulesAbove(parent);
};

type Exported = { readonly types: ReadonlyArray<string>; readonly values: ReadonlyArray<string> };

// The exported names of the package, split into types and values.
const exportedNames = (consumer: string, name: string): Exported => {
  const probe = join(consumer, 'exports-probe.ts');
  writeFileSync(probe, `export * from '${name}';\n`);
  const program = ts.createProgram([probe], { module: ts.ModuleKind.NodeNext, moduleResolution: ts.ModuleResolutionKind.NodeNext, types: [], noEmit: true });
  const checker = program.getTypeChecker();
  const source = program.getSourceFile(probe);
  const moduleSymbol = source === undefined ? undefined : checker.getSymbolAtLocation(source);
  rmSync(probe);
  const exports = moduleSymbol === undefined ? [] : checker.getExportsOfModule(moduleSymbol);
  const resolved = exports.map((symbol) => (symbol.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol));
  const typeFlags = ts.SymbolFlags.TypeAlias | ts.SymbolFlags.Interface | ts.SymbolFlags.Class | ts.SymbolFlags.Enum;
  const named = (flags: number): ReadonlyArray<string> => exports.filter((_, i) => ((resolved[i]?.flags ?? 0) & flags) !== 0).map((symbol) => symbol.name);
  return { types: named(typeFlags), values: named(ts.SymbolFlags.Value) };
};

// The exports that have no wrong use in MISUSES or VALUE_MISUSES.
const untoldExports = ({ types, values }: Exported): ReadonlyArray<string> => [
  ...types.filter((type) => MISUSES[type] === undefined).map((type) => `exported type ${type} has no wrong use in MISUSES (scripts/check-packed-package.ts)`),
  ...values.filter((value) => VALUE_MISUSES[value] === undefined).map((value) => `exported value ${value} has no wrong call in VALUE_MISUSES (scripts/check-packed-package.ts)`),
];

// The consumer file: per exported type, a right use and then a wrong one; per
// exported value, a wrong call. Returns the text and the expected errors, as
// `line:TScode`.
const consumerSource = (name: string, { types, values }: Exported): { readonly text: string; readonly expected: ReadonlyArray<string> } => {
  const lines = [
    ...(types.length === 0 ? [] : [`import type { ${types.join(', ')} } from '${name}';`]),
    ...(values.length === 0 ? [] : [`import { ${values.join(', ')} } from '${name}';`]),
  ];
  const expected: string[] = [];
  for (const type of types) {
    const misuse = MISUSES[type];
    if (misuse === undefined) continue;
    lines.push(`export const fits${type}: ${misuse.use} = ${misuse.fits};`, `export const misuse${type}: ${misuse.use} = ${misuse.misuse};`);
    expected.push(`${lines.length}:TS${misuse.code}`);
  }
  for (const value of values) {
    const misuse = VALUE_MISUSES[value];
    if (misuse === undefined) continue;
    lines.push(`export const misuseValue${value} = ${misuse.misuse};`);
    expected.push(`${lines.length}:TS${misuse.code}`);
  }
  return { text: `${lines.join('\n')}\n`, expected };
};

const typeProblems = (consumer: string, expected: ReadonlyArray<string>): ReadonlyArray<string> => {
  const tsc = run(['bun', TSC, '-p', consumer, '--noEmit', '--skipLibCheck', 'false', '--pretty', 'false'], consumer);
  const reported = [...tsc.out.matchAll(/^uses\.ts\((\d+),\d+\): error (TS\d+)/gm)].map((m) => `${m[1]}:${m[2]}`);
  const elsewhere = tsc.out.split('\n').filter((line) => line.includes('error TS') && !line.startsWith('uses.ts('));
  return [
    ...expected.filter((e) => !reported.includes(e)).map((e) => `consumer typecheck: expected ${e} (line:code), not reported`),
    ...reported.filter((r) => !expected.includes(r)).map((r) => `consumer typecheck: unexpected ${r} (line:code)`),
    ...elsewhere.map((line) => `consumer typecheck: ${line.trim()}`),
  ];
};

const runtimeProblems = (consumer: string, manifest: Manifest): ReadonlyArray<string> => {
  const bin = join(consumer, 'node_modules', manifest.name, withoutDot(Object.values(manifest.bin ?? {})[0] ?? ''));
  const env = isolatedEnv(mkdtempSync(join(consumer, 'home-')));
  return ['node', 'bun'].flatMap((runtime) => {
    const version = run([runtime, bin, '--version'], consumer, env);
    const load = run(
      runtime === 'node' ? ['node', '--input-type=module', '-e', `await import('${manifest.name}')`] : ['bun', '-e', `await import('${manifest.name}')`],
      consumer,
      env
    );
    return [
      ...(version.ok && version.out.trim() === manifest.version ? [] : [`${runtime}: the bin did not answer --version with ${manifest.version}: ${version.out.slice(0, 300)}`]),
      ...(load.ok ? [] : [`${runtime}: the library does not load by name: ${load.out.slice(0, 300)}`]),
    ];
  });
};

const consumerProblems = async (tarball: string, manifest: Manifest): Promise<ReadonlyArray<string>> => {
  const consumer = mkdtempSync(join(tmpdir(), 'packed-consumer-'));
  try {
    if (hasNodeModulesAbove(consumer)) return [`${consumer} has a node_modules above it: the install would not be isolated`];
    writeFileSync(join(consumer, 'package.json'), JSON.stringify({ name: 'packed-consumer', private: true, type: 'module' }));
    writeFileSync(
      join(consumer, 'tsconfig.json'),
      JSON.stringify({ compilerOptions: { module: 'nodenext', moduleResolution: 'nodenext', target: 'es2022', strict: true, types: [] }, files: ['uses.ts'] })
    );
    const install = run(['bun', 'add', tarball], consumer);
    if (!install.ok) return [`the tarball does not install: ${install.out.slice(0, 400)}`];
    const exported = exportedNames(consumer, manifest.name);
    const { text, expected } = consumerSource(manifest.name, exported);
    writeFileSync(join(consumer, 'uses.ts'), text);
    const real = [
      ...untoldExports(exported),
      ...(exported.types.length === 0 ? ['the library exports no type'] : []),
      ...typeProblems(consumer, expected),
      ...runtimeProblems(consumer, manifest),
    ];
    return [...real, ...(await controlProblems(consumer, manifest, exported, expected))];
  } finally {
    rmSync(consumer, { recursive: true, force: true });
  }
};

// The consumer typecheck must fail on a declaration that imports core, and on
// types and values that became `any`. A check that passes these has stopped
// checking.
const controlProblems = async (consumer: string, manifest: Manifest, { types, values }: Exported, expected: ReadonlyArray<string>): Promise<ReadonlyArray<string>> => {
  const index = join(consumer, 'node_modules', manifest.name, 'dist', 'index.d.ts');
  const controls = {
    'an index.d.ts that imports core': `export type { ${types.join(', ')} } from '${CORE}/domain/result.ts';\n`,
    'an index.d.ts whose types and values are any': [
      ...types.map((type) => `export type ${type}<A = any, B = any, C = any, D = any> = any;\n`),
      ...values.map((value) => `export declare const ${value}: any;\n`),
    ].join(''),
  };
  const passed: string[] = [];
  for (const [name, text] of Object.entries(controls)) {
    await Bun.write(index, text);
    if (typeProblems(consumer, expected).length === 0) passed.push(`control: the consumer typecheck passed ${name}`);
  }
  return passed;
};

// The tarball's path comes from stdout alone (--quiet prints only the path).
// stderr goes to the person: a warning there (as `nobody`, an unreadable CA
// bundle) is not the path (.claude/LESSONS.md, 2026-10-18).
const pack = (): { readonly tarball: string; readonly dir: string } => {
  const dir = mkdtempSync(join(tmpdir(), 'packed-'));
  const packed = Bun.spawnSync(['bun', 'pm', 'pack', '--destination', dir, '--quiet', '--ignore-scripts'], { cwd: process.cwd(), stdout: 'pipe', stderr: 'inherit' });
  const out = packed.stdout.toString();
  if (packed.exitCode !== 0) throw new Error(`bun pm pack failed: ${out}`);
  return { tarball: out.trim().split('\n').at(-1) ?? '', dir };
};

const tarballProblems = async (tarball: string, dir: string): Promise<{ readonly problems: ReadonlyArray<string>; readonly manifest: Manifest }> => {
  const listed = run(['tar', '-tzf', tarball], dir);
  const files = listed.out
    .split('\n')
    .filter(Boolean)
    .map((path) => path.replace(/^package\//, ''));
  run(['tar', '-xzf', tarball, '-C', dir], dir);
  const manifest = (await Bun.file(join(dir, 'package', 'package.json')).json()) as Manifest;
  const declarations = Object.fromEntries(
    await Promise.all(files.filter((f) => f.endsWith('.d.ts')).map(async (f) => [f, await Bun.file(join(dir, 'package', f)).text()] as const))
  );
  const bundles = bundleTargets(manifest, existsSync);
  return { problems: [...fileProblems(files, manifest, bundles), ...manifestProblems(manifest), ...declarationProblems(declarations)], manifest };
};

const SELFTEST_MANIFEST: Manifest = {
  name: '@ask-marcel/office-x',
  version: '1.0.0',
  bin: { x: './dist/cli.js' },
  exports: { '.': { types: './dist/index.d.ts', import: './dist/index.js' } },
};
// A package with a helper entry, as auth has: src/main.ts, src/index.ts and src/token.ts exist.
const SELFTEST_BUNDLES = bundleTargets(SELFTEST_MANIFEST, (path) => ['src/main.ts', 'src/index.ts', 'src/token.ts'].includes(path));
const SELFTEST_FILES = [...TOP_FILES, 'dist/cli.js', 'dist/index.js', 'dist/token.js', 'dist/index.d.ts', 'dist/shared/domain/result.d.ts'];
const SELFTEST_EXPORTS: Exported = { types: ['Result'], values: ['resolveAuthPaths'] };
const SELFTEST_CONSUMER = {
  text: [
    "import type { Result } from '@ask-marcel/office-x';",
    "import { resolveAuthPaths } from '@ask-marcel/office-x';",
    'export const fitsResult: Result<number, string> = { ok: true, value: 1 };',
    "export const misuseResult: Result<number, string> = { ok: true, value: 'one' };",
    'export const misuseValueresolveAuthPaths = resolveAuthPaths(1, {});',
    '',
  ].join('\n'),
  expected: ['4:TS2322', '5:TS2345'],
};

const selftest = (): number => {
  const cases: ReadonlyArray<{ readonly name: string; readonly problems: number }> = [
    {
      name: 'a clean package',
      problems: [
        ...fileProblems(SELFTEST_FILES, SELFTEST_MANIFEST, SELFTEST_BUNDLES),
        ...manifestProblems({ ...SELFTEST_MANIFEST, devDependencies: { [CORE]: '1.0.0' } }),
        ...declarationProblems({ 'dist/index.d.ts': "export type { Result } from './shared/domain/result.js';" }),
      ].length,
    },
    { name: 'a stray source file', problems: fileProblems([...SELFTEST_FILES, 'src/index.ts'], SELFTEST_MANIFEST, SELFTEST_BUNDLES).length === 1 ? 0 : 1 },
    {
      name: 'a missing bin',
      problems:
        fileProblems(
          SELFTEST_FILES.filter((f) => f !== 'dist/cli.js'),
          SELFTEST_MANIFEST,
          SELFTEST_BUNDLES
        ).length === 1
          ? 0
          : 1,
    },
    {
      name: 'a missing helper entry',
      problems:
        fileProblems(
          SELFTEST_FILES.filter((f) => f !== 'dist/token.js'),
          SELFTEST_MANIFEST,
          SELFTEST_BUNDLES
        ).length === 1
          ? 0
          : 1,
    },
    {
      name: 'a helper entry in a package with no src/token.ts',
      problems:
        fileProblems(
          SELFTEST_FILES,
          SELFTEST_MANIFEST,
          bundleTargets(SELFTEST_MANIFEST, (path) => path !== 'src/token.ts')
        ).length === 1
          ? 0
          : 1,
    },
    { name: 'core in dependencies', problems: manifestProblems({ ...SELFTEST_MANIFEST, dependencies: { [CORE]: '1.0.0' } }).length === 1 ? 0 : 1 },
    {
      name: 'a declaration that imports core',
      problems: declarationProblems({ 'dist/index.d.ts': `export type { Result } from '${CORE}/domain/result.ts';` }).length === 1 ? 0 : 1,
    },
    {
      name: 'a declaration that reaches ../core',
      problems: declarationProblems({ 'dist/a/b.d.ts': "export type { Result } from '../core/src/domain/result.js';" }).length === 1 ? 0 : 1,
    },
    { name: 'an exported type and value, each with its wrong use', problems: untoldExports(SELFTEST_EXPORTS).length },
    { name: 'an exported type with no wrong use', problems: untoldExports({ ...SELFTEST_EXPORTS, types: ['Result', 'Unlisted'] }).length === 1 ? 0 : 1 },
    { name: 'an exported value with no wrong call', problems: untoldExports({ ...SELFTEST_EXPORTS, values: ['resolveAuthPaths', 'unlisted'] }).length === 1 ? 0 : 1 },
    {
      name: 'a consumer file that imports the values by value and expects an error on each wrong line',
      problems: JSON.stringify(consumerSource('@ask-marcel/office-x', SELFTEST_EXPORTS)) === JSON.stringify(SELFTEST_CONSUMER) ? 0 : 1,
    },
  ];
  const failed = cases.filter((c) => c.problems !== 0);
  for (const c of failed) process.stderr.write(`check-packed-package: SELFTEST FAILED, ${c.name} was judged wrongly\n`);
  if (failed.length === 0)
    process.stdout.write(
      'check-packed-package: selftest passed (a stray file, a missing bin, a missing or stray helper entry, core in dependencies, a declaration naming core and an export with no wrong use are rejected; the consumer file expects an error on each wrong line).\n'
    );
  return failed.length === 0 ? 0 : 1;
};

const main = async (): Promise<number> => {
  if (process.argv.includes('--selftest')) return selftest();
  const { tarball, dir } = pack();
  try {
    const { problems, manifest } = await tarballProblems(tarball, dir);
    const all = [...problems, ...(await consumerProblems(tarball, manifest))];
    for (const problem of all) process.stderr.write(`check-packed-package: ${problem}\n`);
    if (all.length > 0) return 1;
    process.stdout.write(`check-packed-package: ${manifest.name}@${manifest.version} packs clean, and a clean consumer gets the deliberate type errors and nothing else.\n`);
    return 0;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
};

process.exit(await main());
