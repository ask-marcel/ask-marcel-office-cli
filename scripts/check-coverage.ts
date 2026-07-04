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
 *   1  at least one file is below gate, or `bun test` failed
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
  { name: 'infra', prefix: 'src/infra/', threshold: 100 },
  { name: 'composition', prefix: 'src/composition/', threshold: 100 },
  { name: 'presenter', prefix: 'src/presenter/', threshold: 100 },
];

const SKIPPED: ReadonlyArray<SkipRule> = [
  { name: 'test-helpers', match: (p) => p.startsWith('src/test-helpers/') },
  { name: 'entry point', match: (p) => p === 'src/main.ts' },
  // production-wiring: a one-line dynamic-import boundary to the `playwright`
  // runtime dep. Cannot be unit-tested without launching a real browser.
  // Documented at the top of the file itself.
  { name: 'production-wiring', match: (p) => p === 'src/infra/playwright-loader.ts' },
  // public library barrel: re-exports only — no executable logic to cover.
  { name: 'library barrel', match: (p) => p === 'src/index.ts' },
];
// NOTE: src/composition/build-deps.ts USED to be skipped here. It is now
// fully unit-testable via the optional `BuildDepsConfig` argument pattern
// (token store path + logger injectable, sensible defaults preserve prod
// behaviour). See references/architecture.md and references/workflow.md.

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

const parseRow = (line: string): FileRow | undefined => {
  const parts = line.split('|').map((c) => c.trim());
  if (parts.length < 3) return undefined;
  const layouts: ReadonlyArray<{ path: number; funcs: number; lines: number }> = [
    { path: 0, funcs: 1, lines: 2 },
    { path: 1, funcs: 2, lines: 3 },
  ];
  for (const layout of layouts) {
    const path = parts[layout.path];
    if (!path) continue;
    if (path === 'File' || path === 'All files' || path.startsWith('-')) continue;
    if (!path.endsWith('.ts') && !path.endsWith('.tsx')) continue;
    const funcs = Number.parseFloat(parts[layout.funcs] ?? '');
    const lines = Number.parseFloat(parts[layout.lines] ?? '');
    if (Number.isNaN(funcs) || Number.isNaN(lines)) continue;
    const normalised = path.startsWith('./') ? path.slice(2) : path;
    return { path: normalised, funcs, lines };
  }
  return undefined;
};

// We pass `--preload ./scripts/coverage-preload.ts` HERE rather than wiring
// the preload via `bunfig.toml`'s `[test] preload = [...]`. The preload
// side-effect-imports every infra/composition/presenter file (so they show
// up in the coverage table at 0% if untested) but it pulls in heavy
// third-party SDKs (whatever the infra adapters wrap) that add 1–2s to
// every plain `bun test` run. Loading the preload only when computing
// coverage keeps the inner-loop tests fast without losing the gate.
const runTestsWithCoverage = async (): Promise<{ readonly status: number; readonly output: string }> => {
  const proc = Bun.spawn(
    ['bun', 'test', '--coverage', '--preload', './scripts/coverage-preload.ts'],
    { stdout: 'pipe', stderr: 'pipe' }
  );
  const [stdoutText, stderrText] = await Promise.all([
    new Response(proc.stdout).text(),
    new Response(proc.stderr).text(),
  ]);
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

const printViolations = (violations: ReadonlyArray<Violation>): void => {
  console.error('\ncoverage: per-file gate violations:');
  for (const v of violations) {
    console.error(
      `  ${v.file.path}  [${v.tier}]  ${v.metric}=${v.actual.toFixed(1)}%  required=${v.threshold}%`
    );
  }
  const word = violations.length === 1 ? 'violation' : 'violations';
  console.error(
    `\ncoverage: ${violations.length} ${word}. Add tests, or restructure to remove unreachable branches — never lower the threshold.`
  );
};

const main = async (): Promise<number> => {
  const { status, output } = await runTestsWithCoverage();
  if (status !== 0) {
    console.error('\ncoverage: `bun test --coverage` exited non-zero; fix test failures first.');
    return status;
  }
  const rows = output
    .split('\n')
    .map(parseRow)
    .filter((r): r is FileRow => r !== undefined);
  if (rows.length === 0) {
    console.error('\ncoverage: no file rows parsed from the coverage report. Check that `bun test --coverage` is producing a text table.');
    return 1;
  }
  printTierSummary(rows);
  const violations = collectViolations(rows);
  if (violations.length === 0) {
    console.log('\ncoverage: all files meet their tier gate.');
    return 0;
  }
  printViolations(violations);
  return 1;
};

process.exit(await main());
