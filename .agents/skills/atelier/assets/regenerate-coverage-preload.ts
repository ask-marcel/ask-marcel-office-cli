#!/usr/bin/env bun
/*
 * Regenerate scripts/coverage-preload.ts by globbing every TypeScript file
 * under src/infra, src/composition, and src/presenter and emitting a fresh
 * preload that side-effect-imports each one.
 *
 * Why this exists:
 *
 *   `bun test --coverage` only emits rows for files the runner imports.
 *   Untested files in src/infra / src/composition / src/presenter are
 *   silently absent from the coverage table, which makes the per-file gate
 *   trivially pass (a file at 0% is invisible, not failing).
 *
 *   coverage-preload.ts side-effect-imports every such file so they appear
 *   in the table at 0% if untested, which makes the gate fail loudly.
 *
 *   But that file is a manual chore: every new adapter or wiring file means
 *   one more line, and forgetting it silently masks a coverage hole. This
 *   script regenerates the whole preload from the filesystem so the chore
 *   becomes deterministic.
 *
 * Usage:
 *
 *   bun run scripts/regenerate-coverage-preload.ts
 *     # writes scripts/coverage-preload.ts (or `--out <path>` for elsewhere)
 *
 *   bun run scripts/regenerate-coverage-preload.ts --check
 *     # exits non-zero if the on-disk file is out of sync with the glob;
 *     # use this in CI or pre-commit so a missing import blocks the merge.
 *
 *   bun run scripts/regenerate-coverage-preload.ts --selftest
 *     # proves, with Windows path semantics on any machine, that the guard
 *     # below fails and that a Windows path groups and imports with '/'.
 *
 * Either mode exits non-zero, writing and comparing nothing, when the walk
 * found files under a scan directory that the preload would not import.
 *
 * Wire it into pre-commit as an unnumbered pre-flight (the seven fast gates keep their numbers):
 *
 *   echo "[pre-flight] coverage-preload sync" >&2
 *   bun run scripts/regenerate-coverage-preload.ts --check
 *
 * Tune SCAN_DIRS below if your project's Clean Architecture layout differs.
 *
 * See skills/atelier/references/workflow.md (Coverage gates).
 */

import { readdirSync, statSync } from 'node:fs';
import path from 'node:path';

type Args = {
  readonly check: boolean;
  readonly selftest: boolean;
  readonly out: string;
};

const SCAN_DIRS: ReadonlyArray<string> = ['src/infra', 'src/composition', 'src/presenter'];

// Files in scan dirs that should NOT be preloaded. Tests are obvious; ports
// are type-only and have no runtime to preload; index.ts files are usually
// barrel exports already pulled in by their siblings.
const EXCLUDE = (relPath: string): boolean => relPath.endsWith('.test.ts') || relPath.includes('/ports/') || relPath.endsWith('/index.ts');

const parseArgs = (argv: ReadonlyArray<string>): Args => {
  const check = argv.includes('--check');
  const selftest = argv.includes('--selftest');
  const outIdx = argv.indexOf('--out');
  const candidate = outIdx === -1 ? undefined : argv[outIdx + 1];
  const out = candidate ?? 'scripts/coverage-preload.ts';
  return { check, selftest, out };
};

// path.relative answers in the platform's separator (src\infra\x.ts on Windows), and EXCLUDE,
// the grouping and the import lines all read '/', so a path is normalised once, here.
const repoPath = (repoRoot: string, full: string, paths: typeof path = path): string => paths.relative(repoRoot, full).replaceAll('\\', '/');

const walk = (dir: string, repoRoot: string, acc: string[]): void => {
  let entries: ReadonlyArray<string>;
  try {
    entries = readdirSync(dir);
  } catch {
    return; // dir does not exist; harmless
  }
  for (const entry of entries) {
    const full = path.join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) {
      walk(full, repoRoot, acc);
    } else if (st.isFile() && entry.endsWith('.ts')) {
      const rel = repoPath(repoRoot, full);
      if (!EXCLUDE(rel)) acc.push(rel);
    }
  }
};

// Code-unit order, the order a bare sort() gave, so a preload generated before stays in sync.
const byCodeUnit = (a: string, b: string): number => {
  if (a === b) return 0;
  return a < b ? -1 : 1;
};

type Collected = {
  readonly files: ReadonlyArray<string>;
  // How many files the walk kept under each scan directory, counted by the directory it
  // walked, never by reading the path text, so the guard below checks the grouping.
  readonly found: ReadonlyMap<string, number>;
};

const collectFiles = (repoRoot: string): Collected => {
  const found = new Map<string, number>();
  const files: string[] = [];
  for (const scan of SCAN_DIRS) {
    const acc: string[] = [];
    walk(path.join(repoRoot, scan), repoRoot, acc);
    found.set(scan, acc.length);
    files.push(...acc);
  }
  return { files: files.toSorted(byCodeUnit), found };
};

const groupByScanDir = (files: ReadonlyArray<string>): ReadonlyMap<string, ReadonlyArray<string>> =>
  Map.groupBy(files, (f) => SCAN_DIRS.find((d) => f.startsWith(`${d}/`)) ?? 'other');

// The scan directories whose files the walk found but the preload would not import. On
// Windows, before the normalisation above, every path landed in the unwritten 'other' group:
// the preload imported nothing and --check still reported "in sync". An empty directory is
// no failure (the Bun bootstrap creates src/presenter and src/composition empty).
const unemitted = (found: ReadonlyMap<string, number>, grouped: ReadonlyMap<string, ReadonlyArray<string>>): ReadonlyArray<string> =>
  SCAN_DIRS.filter((d) => (found.get(d) ?? 0) > (grouped.get(d)?.length ?? 0));

const printUnemitted = (short: ReadonlyArray<string>, found: ReadonlyMap<string, number>, grouped: ReadonlyMap<string, ReadonlyArray<string>>): void => {
  for (const dir of short) {
    console.error(`coverage-preload: ${dir}/ holds ${found.get(dir) ?? 0} file(s) for the preload, which would import ${grouped.get(dir)?.length ?? 0}.`);
  }
  const stray = grouped.get('other')?.[0];
  if (stray !== undefined) console.error(`  First path outside every scan directory: ${stray}`);
  console.error('  The paths did not group under their scan directory (a path separator, a SCAN_DIRS spelling); nothing was written or compared.');
};

const HEADER = `/*
 * Auto-generated by scripts/regenerate-coverage-preload.ts. Do not hand-edit;
 * run \`bun run scripts/regenerate-coverage-preload.ts\` to rewrite.
 *
 * \`bun test --coverage\` only emits rows for files the runner imports. Untested
 * src/infra, src/composition, and src/presenter files are silently absent
 * from the table, which makes the per-file gate trivially pass. This preload
 * side-effect-imports every such file so they appear at 0% (or better) and
 * the gate can fail loudly.
 *
 * Wired ONLY at coverage time, NOT in bunfig.toml. \`scripts/check-coverage.ts\`
 * spawns \`bun test --coverage --preload ./scripts/coverage-preload.ts\`.
 *
 * See skills/atelier/references/workflow.md.
 */`;

const buildContent = (grouped: ReadonlyMap<string, ReadonlyArray<string>>, repoRoot: string, outPath: string, paths: typeof path = path): string => {
  // Imports are written relative to the output file's directory.
  const outDir = paths.join(repoRoot, outPath, '..');

  const lines: string[] = [HEADER];
  for (const dir of SCAN_DIRS) {
    const list = grouped.get(dir);
    if (!list || list.length === 0) continue;
    lines.push('', `// --- ${dir}/ ---`);
    for (const f of list) {
      // An import specifier is '/' on every platform; path.relative answers '..\src\..' on Windows.
      const fromOut = paths.relative(outDir, paths.join(repoRoot, f)).replaceAll('\\', '/');
      const importPath = fromOut.startsWith('.') ? fromOut : `./${fromOut}`;
      lines.push(`import '${importPath}';`);
    }
  }
  lines.push(''); // trailing newline
  return lines.join('\n');
};

// The generator proves it can fail before it writes (canon 15.10), with Windows path semantics
// (path.win32) on any machine: the shape path.relative gives there trips the guard unnormalised,
// and once normalised it groups, imports with '/', and EXCLUDE still reads it.
const selftest = (): number => {
  const root = String.raw`C:\repo`;
  const unnormalised = path.win32.relative(root, String.raw`C:\repo\src\infra\http.ts`);
  const normalised = repoPath(root, String.raw`C:\repo\src\infra\http.ts`, path.win32);
  const found = new Map([['src/infra', 1]]);
  const grouped = groupByScanDir([normalised]);
  const content = buildContent(grouped, root, 'scripts/coverage-preload.ts', path.win32);
  const failures = [
    unemitted(found, groupByScanDir([unnormalised])).join(', ') === 'src/infra' ? '' : `the unnormalised Windows path ${unnormalised} left the guard quiet`,
    unemitted(found, grouped).length === 0 ? '' : `the normalised Windows path ${normalised} did not group under src/infra`,
    content.includes("import '../src/infra/http.ts';") ? '' : 'the Windows import line is not ../src/infra/http.ts',
    EXCLUDE(repoPath(root, String.raw`C:\repo\src\infra\ports\clock.ts`, path.win32)) ? '' : 'a Windows path under ports/ was not excluded',
  ].filter((f) => f !== '');
  for (const failure of failures) {
    console.error(`coverage-preload: selftest FAIL: ${failure}`);
  }
  if (failures.length > 0) return 1;
  console.log(
    'coverage-preload: selftest OK: an unnormalised Windows path fails the run, a normalised one groups under its scan directory and imports with /, ports/ stays excluded'
  );
  return 0;
};

const main = async (): Promise<number> => {
  const args = parseArgs(process.argv.slice(2));
  if (args.selftest) return selftest();
  const repoRoot = process.cwd();
  const { files, found } = collectFiles(repoRoot);
  const grouped = groupByScanDir(files);
  const short = unemitted(found, grouped);
  if (short.length > 0) {
    printUnemitted(short, found, grouped);
    return 1;
  }
  const content = buildContent(grouped, repoRoot, args.out);

  if (args.check) {
    const outFile = Bun.file(args.out);
    const existing = (await outFile.exists()) ? await outFile.text() : '';
    if (existing.trim() === content.trim()) {
      console.log(`coverage-preload: in sync (${files.length} files)`);
      return 0;
    }
    console.error('coverage-preload: OUT OF SYNC.');
    console.error(`  ${args.out} does not match the current set of source files.`);
    console.error('  Run: bun run scripts/regenerate-coverage-preload.ts');
    console.error(`  Then: git add ${args.out} && commit.`);
    return 1;
  }

  await Bun.write(args.out, content);
  console.log(`coverage-preload: wrote ${args.out} (${files.length} files)`);
  return 0;
};

process.exit(await main());
