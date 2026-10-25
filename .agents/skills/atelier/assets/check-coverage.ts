#!/usr/bin/env bun
/*
 * Per-tier coverage gate for atelier Clean Architecture.
 *
 * Runs `bun test --coverage`, parses the text report, and enforces a
 * different threshold for each source tier (domain, use-cases, infra,
 * composition, presenter). Exits non-zero on any violation.
 *
 * Exit codes:
 *   0  every non-skipped file meets its tier's threshold
 *   1  at least one file is below gate, a file under src/ matches no tier and
 *      no skip rule, or `bun test` failed
 *
 * `--selftest` runs planted table rows through the same parser and checks,
 * without `bun test`, and exits 1 unless each planted violation is caught.
 *
 * Tune per-project by editing COVERAGE_RULES and SKIPPED below.
 *
 * IMPORTANT: bunfig.toml MUST NOT set a global `coverageThreshold` when
 * this script owns enforcement. The global threshold would make
 * `bun test --coverage` exit non-zero before the script can parse,
 * and the per-file violation breakdown never prints. See
 * skills/atelier/references/workflow.md.
 */

type Tier = {
  readonly name: string;
  readonly prefix: string;
  readonly threshold: number;
};

type SkipRule = {
  readonly name: string;
  readonly match: (path: string) => boolean;
};

const COVERAGE_RULES: ReadonlyArray<Tier> = [
  { name: 'domain', prefix: 'src/domain/', threshold: 100 },
  { name: 'use-cases', prefix: 'src/use-cases/', threshold: 100 },
  { name: 'infra', prefix: 'src/infra/', threshold: 80 },
  { name: 'composition', prefix: 'src/composition/', threshold: 80 },
  { name: 'presenter', prefix: 'src/presenter/', threshold: 80 },
];

const SKIPPED: ReadonlyArray<SkipRule> = [
  { name: 'test-helpers', match: (p) => p.startsWith('src/test-helpers/') },
  { name: 'entry point', match: (p) => p === 'src/main.ts' },
];
// Resist adding composition/wiring files here: any composition root is
// 100%-testable once its state-sources (paths, env, clock) are parameters
// and its sinks (logger, sender) injected: see references/architecture.md
// (Composition root testability). SKIPPED is for genuine non-code entries.

type FileRow = {
  readonly path: string;
  readonly funcs: number;
  readonly lines: number;
};

type Violation = {
  readonly file: FileRow;
  readonly tier: string;
  readonly threshold: number;
  readonly metric: 'funcs' | 'lines';
  readonly actual: number;
};

const isSkipped = (path: string): boolean => SKIPPED.some((s) => s.match(path));

const findTier = (path: string): Tier | undefined => COVERAGE_RULES.find((t) => path.startsWith(t.prefix));

// Bun prints the table with or without a leading empty column depending on the
// version, hence two column layouts tried in order.
type Layout = { readonly path: number; readonly funcs: number; readonly lines: number };

const LAYOUTS: ReadonlyArray<Layout> = [
  { path: 0, funcs: 1, lines: 2 },
  { path: 1, funcs: 2, lines: 3 },
];

const isSourcePath = (path: string): boolean => path !== 'File' && path !== 'All files' && !path.startsWith('-') && (path.endsWith('.ts') || path.endsWith('.tsx'));

const percentAt = (parts: ReadonlyArray<string>, index: number): number => Number.parseFloat(parts[index] ?? '');

const rowAt = (parts: ReadonlyArray<string>, layout: Layout): FileRow | undefined => {
  const raw = parts[layout.path];
  if (!raw || !isSourcePath(raw)) return undefined;
  const funcs = percentAt(parts, layout.funcs);
  const lines = percentAt(parts, layout.lines);
  if (Number.isNaN(funcs) || Number.isNaN(lines)) return undefined;
  // Bun prints the platform's separator (src\domain\x.ts on Windows), and every prefix
  // and skip rule here is written with '/', so the path is normalised once, where it enters.
  const path = raw.replaceAll('\\', '/');
  return { path: path.startsWith('./') ? path.slice(2) : path, funcs, lines };
};

const parseRow = (line: string): FileRow | undefined => {
  const parts = line.split('|').map((c) => c.trim());
  if (parts.length < 3) return undefined;
  for (const layout of LAYOUTS) {
    const row = rowAt(parts, layout);
    if (row) return row;
  }
  return undefined;
};

const parseRows = (output: string): ReadonlyArray<FileRow> =>
  output
    .split('\n')
    .map((line) => parseRow(line))
    .filter((r): r is FileRow => r !== undefined);

// We pass `--preload ./scripts/coverage-preload.ts` HERE rather than wiring
// the preload via `bunfig.toml`'s `[test] preload = [...]`. The preload
// side-effect-imports every infra/composition/presenter file (so they show
// up in the coverage table at 0% if untested) but it pulls in heavy
// third-party SDKs (whatever the infra adapters wrap) that add 1–2s to
// every plain `bun test` run. Loading the preload only when computing
// coverage keeps the inner-loop tests fast without losing the gate.
const runTestsWithCoverage = async (): Promise<{ readonly status: number; readonly output: string }> => {
  const proc = Bun.spawn(['bun', 'test', '--coverage', '--preload', './scripts/coverage-preload.ts'], { stdout: 'pipe', stderr: 'pipe' });
  const [stdoutText, stderrText] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text()]);
  process.stdout.write(stdoutText);
  process.stderr.write(stderrText);
  const status = await proc.exited;
  return { status, output: `${stdoutText}\n${stderrText}` };
};

const worstBy = (rows: ReadonlyArray<FileRow>, metric: 'funcs' | 'lines'): FileRow | undefined => {
  const [first, ...rest] = rows;
  if (!first) return undefined;
  return rest.reduce((w, r) => (r[metric] < w[metric] ? r : w), first);
};

const printTierSummary = (rows: ReadonlyArray<FileRow>): void => {
  console.log('\ncoverage: tier summary (worst funcs / worst lines):');
  for (const tier of COVERAGE_RULES) {
    const inTier = rows.filter((r) => !isSkipped(r.path) && r.path.startsWith(tier.prefix));
    if (inTier.length === 0) {
      console.log(`  ${tier.name.padEnd(12)} (>= ${tier.threshold}%)  no files`);
      continue;
    }
    const worstFuncs = worstBy(inTier, 'funcs');
    const worstLines = worstBy(inTier, 'lines');
    if (!worstFuncs || !worstLines) continue;
    console.log(
      `  ${tier.name.padEnd(12)} (>= ${tier.threshold}%)  funcs: ${worstFuncs.funcs.toFixed(1)}% (${worstFuncs.path})  lines: ${worstLines.lines.toFixed(1)}% (${worstLines.path})`
    );
  }
};

const collectViolations = (rows: ReadonlyArray<FileRow>): ReadonlyArray<Violation> => {
  const violations: Violation[] = [];
  for (const row of rows) {
    if (isSkipped(row.path)) continue;
    const tier = findTier(row.path);
    if (!tier) continue;
    if (row.funcs < tier.threshold) {
      violations.push({ file: row, tier: tier.name, threshold: tier.threshold, metric: 'funcs', actual: row.funcs });
    }
    if (row.lines < tier.threshold) {
      violations.push({ file: row, tier: tier.name, threshold: tier.threshold, metric: 'lines', actual: row.lines });
    }
  }
  return violations;
};

// A row under src/ that no tier and no skip rule claims is judged by nothing, so it fails
// the run by name: a new directory, or a path the prefixes cannot read, would otherwise pass
// in silence (the Windows table did, with "no files" in every tier).
const collectUnmatched = (rows: ReadonlyArray<FileRow>): ReadonlyArray<FileRow> =>
  rows.filter((r) => r.path.startsWith('src/') && !isSkipped(r.path) && findTier(r.path) === undefined);

const printUnmatched = (unmatched: ReadonlyArray<FileRow>): void => {
  console.error('\ncoverage: files under src/ that no tier and no skip rule claims:');
  for (const row of unmatched) {
    console.error(`  ${row.path}`);
  }
  console.error('\ncoverage: add the directory to COVERAGE_RULES (or a genuine non-code entry to SKIPPED); a file no gate judges passes unchecked.');
};

const printViolations = (violations: ReadonlyArray<Violation>): void => {
  console.error('\ncoverage: per-file gate violations:');
  for (const v of violations) {
    console.error(`  ${v.file.path}  [${v.tier}]  ${v.metric}=${v.actual.toFixed(1)}%  required=${v.threshold}%`);
  }
  const word = violations.length === 1 ? 'violation' : 'violations';
  console.error(`\ncoverage: ${violations.length} ${word}. Add tests, or restructure to remove unreachable branches, never lower the threshold.`);
};

const main = async (): Promise<number> => {
  const { status, output } = await runTestsWithCoverage();
  if (status !== 0) {
    console.error('\ncoverage: `bun test --coverage` exited non-zero; fix test failures first.');
    return status;
  }
  const rows = parseRows(output);
  if (rows.length === 0) {
    console.error('\ncoverage: no file rows parsed from the coverage report. Check that `bun test --coverage` is producing a text table.');
    return 1;
  }
  printTierSummary(rows);
  const violations = collectViolations(rows);
  const unmatched = collectUnmatched(rows);
  if (violations.length === 0 && unmatched.length === 0) {
    console.log('\ncoverage: all files meet their tier gate.');
    return 0;
  }
  if (violations.length > 0) printViolations(violations);
  if (unmatched.length > 0) printUnmatched(unmatched);
  return 1;
};

// The gate proves it can fail before it judges (canon 15.10): planted rows go through the
// parser and checks a real run uses. Bun on Windows prints backslashes; both layouts appear.
const SELFTEST_TABLE: ReadonlyArray<string> = [
  'File                         | % Funcs | % Lines | Uncovered Line #s',
  'All files                    |   72.00 |   70.00 |',
  String.raw` src\domain\shout.ts          |   50.00 |   40.00 | 3-4`,
  ' src/infra/http.ts            |   90.00 |   85.00 |',
  '| src/jobs/tick.ts | 100.00 | 100.00 | |',
  String.raw` src\test-helpers\fake.ts     |    0.00 |    0.00 | 1-9`,
  ' src/main.ts                  |    0.00 |    0.00 | 1-3',
  ' scripts/coverage-preload.ts  |  100.00 |  100.00 |',
];

const selftest = (): number => {
  const rows = parseRows(SELFTEST_TABLE.join('\n'));
  const caught = collectViolations(rows).map((v) => `${v.file.path} ${v.tier} ${v.metric}`);
  const unmatched = collectUnmatched(rows).map((r) => r.path);
  const failures = [
    caught.join(', ') === 'src/domain/shout.ts domain funcs, src/domain/shout.ts domain lines'
      ? ''
      : `a Windows row under the domain gate gave [${caught.join(', ')}], not its funcs and lines violations`,
    unmatched.join(', ') === 'src/jobs/tick.ts' ? '' : `the src/ row in no tier gave [${unmatched.join(', ')}], not src/jobs/tick.ts alone`,
  ].filter((f) => f !== '');
  for (const failure of failures) {
    console.error(`coverage: selftest FAIL: ${failure}`);
  }
  if (failures.length > 0) return 1;
  console.log(
    'coverage: selftest OK: a Windows row under its tier gate fails on funcs and lines, a src/ row in no tier and no skip rule fails by name, skipped and non-src rows pass, both table layouts parse'
  );
  return 0;
};

process.exit(process.argv.includes('--selftest') ? selftest() : await main());
