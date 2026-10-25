#!/usr/bin/env bun
/*
 * Per-file mutation gate for the Stryker runs (from the atelier asset
 * check-mutation-score.ts, adapted to the workspace folders and the shards).
 *
 * Stryker's `thresholds.break` judges the score of the whole run, so a weak
 * file passes beside strong ones (atelier, 2026-10-23: a file at 55.56 passed
 * in a run that totalled 90.00). This reads the JSON report that the run just
 * wrote and holds each file to the same `break`, read from the report.
 *
 * Run it from a workspace folder (core, auth, read, write), where the report
 * is, after a Stryker run. scripts/mutate-changed.sh and
 * scripts/mutate-staged.sh call it with the files that the run mutated; a CI
 * shard thus checks its own files only.
 *
 *   bun run ../scripts/check-mutation-score.ts                  # every file in the report (a full sweep)
 *   bun run ../scripts/check-mutation-score.ts src/domain/a.ts  # only the files this run mutated
 *   bun run scripts/check-mutation-score.ts --selftest          # planted reports, no Stryker
 *
 * Adaptation: Stryker writes a file into the report only when the file has a
 * mutant, so a changed file with no mutant (types only, such as
 * core/src/use-cases/commands/command-types.ts) is not in it. A named file is
 * thus judged so:
 *   - in the report: its score must be at `break` or more;
 *   - not in the report, but in the `--mutate` list of the run (the report
 *     keeps it under `config.mutate`): the run had no mutant for it, a pass
 *     with a note;
 *   - in neither: the run did not mutate it (an old report, or a run on other
 *     files). It fails by name, so a run that mutated nothing cannot pass.
 *
 * Exit codes:
 *   0  every file in scope scores at or above `break`
 *   1  a file is below `break` or not mutated, or the report is missing,
 *      unreadable or has no `break`
 */

type Mutant = { readonly status: string };
type ReportFile = { readonly mutants: ReadonlyArray<Mutant> };
type Report = {
  readonly files: Readonly<Record<string, ReportFile>>;
  readonly thresholds?: { readonly break?: number | null };
  readonly config?: { readonly mutate?: ReadonlyArray<string> | string };
};

type FileScore = { readonly path: string; readonly score: number | undefined; readonly state: 'scored' | 'no-mutant' | 'missing' };
type Verdict = { readonly lines: ReadonlyArray<string>; readonly failures: ReadonlyArray<string> };

const REPORT_PATH = 'reports/mutation/mutation.json';

// Stryker's own score: detected (killed or timed out) over valid (detected,
// survived or not covered). Compile and runtime errors, ignored and pending
// mutants count for neither.
const DETECTED: ReadonlySet<string> = new Set(['Killed', 'Timeout']);
const UNDETECTED: ReadonlySet<string> = new Set(['Survived', 'NoCoverage']);

const scoreOf = (file: ReportFile): number | undefined => {
  const detected = file.mutants.filter((m) => DETECTED.has(m.status)).length;
  const valid = detected + file.mutants.filter((m) => UNDETECTED.has(m.status)).length;
  return valid === 0 ? undefined : (detected / valid) * 100;
};

// Report keys and git paths are compared with '/', whatever the platform printed.
const normalise = (path: string): string => {
  const slashed = path.replaceAll('\\', '/');
  return slashed.startsWith('./') ? slashed.slice(2) : slashed;
};

// The --mutate value of the run: a list in the report, or one comma-separated string.
const mutatedByRun = (report: Report): ReadonlySet<string> => {
  const mutate = report.config?.mutate ?? [];
  const entries = typeof mutate === 'string' ? mutate.split(',') : mutate;
  return new Set(entries.map((entry) => normalise(entry.trim())));
};

const scoreFiles = (report: Report, scope: ReadonlyArray<string>): ReadonlyArray<FileScore> => {
  const files = new Map(Object.entries(report.files).map(([path, file]) => [normalise(path), file]));
  const inRun = mutatedByRun(report);
  const paths = scope.length > 0 ? scope.map((p) => normalise(p)) : [...files.keys()];
  return paths.map((path): FileScore => {
    const file = files.get(path);
    if (file) return { path, score: scoreOf(file), state: 'scored' };
    return { path, score: undefined, state: inRun.has(path) ? 'no-mutant' : 'missing' };
  });
};

const lineOf = (s: FileScore): string | undefined => {
  if (s.state === 'missing') return undefined;
  if (s.state === 'no-mutant') return `  ${s.path}  no mutant (the run mutated it and found nothing to mutate)`;
  return `  ${s.path}  ${s.score === undefined ? 'no valid mutant' : s.score.toFixed(2)}`;
};

const failureOf = (s: FileScore, threshold: number): string | undefined => {
  if (s.state === 'missing') return `${s.path}  not in the report (the run did not mutate it)`;
  if (s.score !== undefined && s.score < threshold) return `${s.path}  score=${s.score.toFixed(2)}  required=${threshold}`;
  return undefined;
};

const judge = (report: Report, scope: ReadonlyArray<string>): Verdict => {
  const threshold = report.thresholds?.break;
  if (typeof threshold !== 'number') {
    return { lines: [], failures: ['the report carries no thresholds.break; set it in stryker.conf.json (atelier: 90)'] };
  }
  const scores = scoreFiles(report, scope);
  return {
    lines: scores.map(lineOf).filter((l): l is string => l !== undefined),
    failures: scores.map((s) => failureOf(s, threshold)).filter((f): f is string => f !== undefined),
  };
};

const readReport = async (path: string): Promise<Report | undefined> => {
  const file = Bun.file(path);
  if (!(await file.exists())) return undefined;
  try {
    return (await file.json()) as Report;
  } catch {
    return undefined;
  }
};

const main = async (scope: ReadonlyArray<string>): Promise<number> => {
  const report = await readReport(REPORT_PATH);
  if (!report) {
    console.error(`mutation: no readable ${REPORT_PATH}; the json reporter must be on in stryker.conf.json and the run must have finished.`);
    return 1;
  }
  const { lines, failures } = judge(report, scope);
  console.log('mutation: per-file scores');
  for (const line of lines) console.log(line);
  if (failures.length === 0) {
    console.log('mutation: every file meets the break threshold.');
    return 0;
  }
  console.error('\nmutation: per-file gate violations:');
  for (const failure of failures) console.error(`  ${failure}`);
  console.error('\nmutation: strengthen the tests that pin these files, never lower the threshold.');
  return 1;
};

// The gate proves it can fail before it judges (canon 15.10): planted reports
// go through the same scoring and checks that a real run uses.
const mutants = (statuses: ReadonlyArray<string>): ReportFile => ({ mutants: statuses.map((status) => ({ status })) });

const SELFTEST_REPORT: Report = {
  thresholds: { break: 90 },
  config: { mutate: ['src/domain/strong.ts', 'src/domain/weak.ts', 'src/use-cases/types-only.ts', 'src/use-cases/no-mutant.ts'] },
  files: {
    // 18 of 20 detected: 90.00, at the line
    'src/domain/strong.ts': mutants([...Array.from({ length: 17 }, () => 'Killed'), 'Timeout', 'Survived', 'NoCoverage']),
    // 5 of 9: 55.56, the file the aggregate hid
    [String.raw`src\domain\weak.ts`]: mutants(['Killed', 'Killed', 'Killed', 'Killed', 'Killed', 'Survived', 'Survived', 'Survived', 'Survived']),
    // errors and ignored mutants count for neither side
    'src/use-cases/types-only.ts': mutants(['CompileError', 'Ignored']),
    'src/domain/stale.ts': mutants(['Survived']),
  },
};

const selftest = (): number => {
  const scoped = judge(SELFTEST_REPORT, ['src/domain/strong.ts', 'src/domain/weak.ts', 'src/use-cases/types-only.ts', 'src/use-cases/no-mutant.ts', 'src/domain/new.ts']);
  const whole = judge(SELFTEST_REPORT, []);
  const unset = judge({ files: SELFTEST_REPORT.files }, []);
  const csv = judge({ ...SELFTEST_REPORT, config: { mutate: 'src/domain/strong.ts,src/use-cases/no-mutant.ts' } }, ['src/use-cases/no-mutant.ts']);
  const checks: ReadonlyArray<readonly [boolean, string]> = [
    [
      scoped.failures.join(' | ') === 'src/domain/weak.ts  score=55.56  required=90 | src/domain/new.ts  not in the report (the run did not mutate it)',
      `a scoped run gave [${scoped.failures.join(' | ')}], not the weak file and the missing one`,
    ],
    [scoped.lines.some((l) => l.includes('src/use-cases/no-mutant.ts  no mutant')), 'a mutated file with no mutant was not noted as such'],
    [whole.failures.length === 2 && whole.failures.some((f) => f.startsWith('src/domain/stale.ts')), `the whole report gave [${whole.failures.join(' | ')}], not weak and stale`],
    [unset.failures.length === 1 && unset.failures[0]?.includes('thresholds.break') === true, `a report with no break gave [${unset.failures.join(' | ')}]`],
    [csv.failures.length === 0, `a --mutate string with the file gave [${csv.failures.join(' | ')}]`],
  ];
  const failed = checks.filter(([ok]) => !ok).map(([, message]) => message);
  for (const message of failed) console.error(`mutation: selftest FAIL: ${message}`);
  if (failed.length > 0) return 1;
  console.log(
    'mutation: selftest OK: a file under break fails beside one at it, a scoped file the run did not mutate fails, a mutated file with no mutant passes with a note, an old file outside the scope is ignored, a Windows key matches, a report with no break fails'
  );
  return 0;
};

const args = process.argv.slice(2);
process.exit(args.includes('--selftest') ? selftest() : await main(args));
