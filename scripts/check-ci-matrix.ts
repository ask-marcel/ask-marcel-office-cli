#!/usr/bin/env bun
/*
 * The CI matrix covers every workspace folder (package split, Checks and CI).
 *
 * Each workspace folder (core, auth, read, write) runs its own lint, tests,
 * coverage and mutation in its own job of the `workspace` matrix in
 * .github/workflows/ci.yml. A folder that the root package.json lists but the
 * matrix forgets has no gate at all, and nothing else reports it: its sources
 * go unmutated, which the mutation scope cannot see from inside the folder.
 * So the matrix's folders must equal the workspaces, both ways.
 *
 * Read's mutation is the one exception (package split, phase 6): its scope is
 * too big for one job, so the `mutation` job runs it in read/, in parallel
 * shards, and the mutation step of the workspace job skips read. So the
 * `mutation` job must run in read with a SHARD_COUNT equal to its shard list
 * (else some files go unmutated), and the workspace step must skip read and
 * only read (else read mutates twice, or another folder not at all).
 *
 * Each published folder checks its command reference after its build (package
 * split, phase 7). The workspace job runs `bun run check:docs` after `bun run
 * build`, for the published folders. The `check:docs` script of each published
 * folder names `gen-docs.ts`, `--check` and `--commands-doc COMMANDS.md`. Else
 * the freshness gate can go dark, and a stale COMMANDS.md passes CI.
 *
 *   --selftest  proves a missing folder, an extra folder, a missing matrix, a
 *               missing or misplaced mutation job, a wrong shard count and a
 *               wrong skip are rejected, and a right workflow passes (rule 15.10).
 *               It also proves that no docs step, a docs step before the build,
 *               a `check:docs` script without `--check` and a published folder
 *               with no `check:docs` script are rejected.
 *
 * Exit codes: 0 the matrix lists exactly the workspaces, read mutates in its
 * shards only, and each published folder checks its command reference; 1 it
 * does not.
 */

type Step = { readonly name?: string; readonly if?: string; readonly run?: string; readonly env?: Readonly<Record<string, unknown>> };
type Job = {
  readonly defaults?: { readonly run?: { readonly 'working-directory'?: string } };
  readonly strategy?: {
    readonly matrix?: { readonly include?: ReadonlyArray<{ readonly folder?: string; readonly published?: boolean }>; readonly shard?: ReadonlyArray<unknown> };
  };
  readonly steps?: ReadonlyArray<Step>;
};
type Workflow = { readonly jobs?: Readonly<Record<string, Job>> };

const JOB = 'workspace';
const MUTATION_JOB = 'mutation';
// The folder whose mutation runs in the shards of the mutation job, not in the workspace job.
const SHARDED = 'read';
const SKIP_SHARDED = `matrix.folder != '${SHARDED}'`;

const matrixFolders = (workflow: Workflow): ReadonlyArray<string> | undefined => {
  const include = workflow.jobs?.[JOB]?.strategy?.matrix?.include;
  return include?.map((entry) => entry.folder ?? '');
};

const matrixProblems = (workspaces: ReadonlyArray<string>, workflow: Workflow): ReadonlyArray<string> => {
  const folders = matrixFolders(workflow);
  if (folders === undefined) return [`ci.yml has no jobs.${JOB}.strategy.matrix.include`];
  return [
    ...workspaces.filter((w) => !folders.includes(w)).map((w) => `workspace folder ${w} has no job in the ${JOB} matrix`),
    ...folders.filter((f) => !workspaces.includes(f)).map((f) => `the ${JOB} matrix names ${f || '(no folder)'}, which is not a workspace`),
  ];
};

// The mutation job runs in the sharded folder, and its SHARD_COUNT equals its shard list.
const mutationJobProblems = (workflow: Workflow): ReadonlyArray<string> => {
  const job = workflow.jobs?.[MUTATION_JOB];
  if (job === undefined) return [`ci.yml has no ${MUTATION_JOB} job, so ${SHARDED} goes unmutated`];
  const folder = job.defaults?.run?.['working-directory'];
  const shards = job.strategy?.matrix?.shard ?? [];
  const counts = (job.steps ?? []).flatMap((step) => (step.env?.['SHARD_COUNT'] === undefined ? [] : [Number(step.env['SHARD_COUNT'])]));
  return [
    ...(folder === SHARDED ? [] : [`the ${MUTATION_JOB} job runs in ${folder ?? 'the root'}, not in ${SHARDED}`]),
    ...(shards.length > 0 && counts.length > 0 && counts.every((count) => count === shards.length)
      ? []
      : [`the ${MUTATION_JOB} job has ${shards.length} shard(s), but its SHARD_COUNT is ${counts.join(', ') || 'missing'}`]),
  ];
};

// The mutation step of the workspace job skips the sharded folder and only it.
const workspaceMutationProblems = (workflow: Workflow): ReadonlyArray<string> => {
  const step = workflow.jobs?.[JOB]?.steps?.find((candidate) => (candidate.name ?? '').startsWith('mutation'));
  if (step === undefined) return [`the ${JOB} job has no mutation step`];
  return step.if === SKIP_SHARDED
    ? []
    : [`the mutation step of the ${JOB} job must skip ${SHARDED} and only ${SHARDED} (if: ${SKIP_SHARDED}), not ${step.if ?? 'run for every folder'}`];
};

const BUILD = 'bun run build';
const CHECK_DOCS = 'bun run check:docs';
const PUBLISHED = 'matrix.published';
// The words of the check:docs script of a published folder: the freshness check of gen-docs on its COMMANDS.md.
const CHECK_DOCS_WORDS = ['gen-docs.ts', '--check', '--commands-doc COMMANDS.md'];

const publishedFolders = (workflow: Workflow): ReadonlyArray<string> =>
  (workflow.jobs?.[JOB]?.strategy?.matrix?.include ?? []).flatMap((entry) => (entry.published === true && entry.folder !== undefined ? [entry.folder] : []));

// The workspace job checks the docs after the build, and the check:docs script of each published folder runs the freshness check.
const docsProblems = (workflow: Workflow, docsScripts: Readonly<Record<string, string | undefined>>): ReadonlyArray<string> => {
  const steps = workflow.jobs?.[JOB]?.steps ?? [];
  const build = steps.findIndex((step) => step.run === BUILD && step.if === PUBLISHED);
  const docs = steps.findIndex((step) => step.run === CHECK_DOCS && step.if === PUBLISHED);
  return [
    ...(build !== -1 && docs > build ? [] : [`the ${JOB} job must run \`${CHECK_DOCS}\` after \`${BUILD}\` (if: ${PUBLISHED})`]),
    ...publishedFolders(workflow).flatMap((folder) => {
      const missing = CHECK_DOCS_WORDS.filter((word) => !(docsScripts[folder] ?? '').includes(word));
      return missing.length === 0 ? [] : [`the check:docs script of ${folder} must name ${missing.join(', ')}`];
    }),
  ];
};

const problems = (workspaces: ReadonlyArray<string>, workflow: Workflow): ReadonlyArray<string> => [
  ...matrixProblems(workspaces, workflow),
  ...mutationJobProblems(workflow),
  ...workspaceMutationProblems(workflow),
];

const selftest = (): number => {
  const mutationJob = (folder: string | undefined, shards: number, count: number): Job => ({
    ...(folder === undefined ? {} : { defaults: { run: { 'working-directory': folder } } }),
    strategy: { matrix: { shard: Array.from({ length: shards }, (_, index) => index) } },
    steps: [{ name: 'mutation', env: { SHARD_COUNT: count } }],
  });
  const workflow = (folders: ReadonlyArray<string> | undefined, mutation: Job | undefined, skip: string | undefined): Workflow => ({
    jobs: {
      [JOB]: {
        ...(folders === undefined ? {} : { strategy: { matrix: { include: folders.map((folder) => ({ folder })) } } }),
        steps: [{ name: 'mutation (changed files)', ...(skip === undefined ? {} : { if: skip }) }],
      },
      ...(mutation === undefined ? {} : { [MUTATION_JOB]: mutation }),
    },
  });
  const sharded = mutationJob(SHARDED, 8, 8);
  const docsWorkflow = (steps: ReadonlyArray<Step>): Workflow => ({
    jobs: {
      [JOB]: {
        strategy: {
          matrix: {
            include: [
              { folder: 'core', published: false },
              { folder: 'read', published: true },
            ],
          },
        },
        steps,
      },
    },
  });
  const build = { run: BUILD, if: PUBLISHED };
  const docs = { run: CHECK_DOCS, if: PUBLISHED };
  const script = 'bun run ../scripts/gen-docs.ts --package src/composition/single-package.ts --check --commands-doc COMMANDS.md';
  const cases = [
    { name: 'a right workflow', found: problems(['core', 'auth'], workflow(['auth', 'core'], sharded, SKIP_SHARDED)).length, expected: 0 },
    { name: 'a missing folder', found: problems(['core', 'auth', 'read'], workflow(['core', 'auth'], sharded, SKIP_SHARDED)).length, expected: 1 },
    { name: 'an extra folder', found: problems(['core'], workflow(['core', 'read'], sharded, SKIP_SHARDED)).length, expected: 1 },
    { name: 'no matrix', found: problems(['core'], workflow(undefined, sharded, SKIP_SHARDED)).length, expected: 1 },
    { name: 'no mutation job', found: problems(['core'], workflow(['core'], undefined, SKIP_SHARDED)).length, expected: 1 },
    { name: 'a mutation job at the root', found: problems(['core'], workflow(['core'], mutationJob(undefined, 8, 8), SKIP_SHARDED)).length, expected: 1 },
    { name: 'a wrong shard count', found: problems(['core'], workflow(['core'], mutationJob(SHARDED, 8, 4), SKIP_SHARDED)).length, expected: 1 },
    { name: 'an unskipped read', found: problems(['core'], workflow(['core'], sharded, undefined)).length, expected: 1 },
    { name: 'a skip of another folder', found: problems(['core'], workflow(['core'], sharded, "matrix.folder != 'write'")).length, expected: 1 },
    { name: 'a right docs step and script', found: docsProblems(docsWorkflow([build, docs]), { read: script }).length, expected: 0 },
    { name: 'no docs step', found: docsProblems(docsWorkflow([build]), { read: script }).length, expected: 1 },
    { name: 'a docs step before the build', found: docsProblems(docsWorkflow([docs, build]), { read: script }).length, expected: 1 },
    { name: 'a check:docs script without --check', found: docsProblems(docsWorkflow([build, docs]), { read: script.replace(' --check', '') }).length, expected: 1 },
    { name: 'a published folder with no check:docs script', found: docsProblems(docsWorkflow([build, docs]), {}).length, expected: 1 },
  ];
  const wrong = cases.filter((c) => c.found !== c.expected);
  for (const c of wrong) process.stderr.write(`check-ci-matrix: SELFTEST FAILED, ${c.name} gave ${c.found} problem(s), expected ${c.expected}\n`);
  if (wrong.length === 0)
    process.stdout.write(
      'check-ci-matrix: selftest passed (a missing folder, an extra folder, no matrix, no or a misplaced mutation job, a wrong shard count, a wrong skip, no docs step, a docs step before the build, a check:docs script without --check and a missing check:docs script are rejected).\n'
    );
  return wrong.length === 0 ? 0 : 1;
};

const main = async (): Promise<number> => {
  if (process.argv.includes('--selftest')) return selftest();
  const { workspaces } = (await Bun.file('package.json').json()) as { readonly workspaces: ReadonlyArray<string> };
  const workflow = Bun.YAML.parse(await Bun.file('.github/workflows/ci.yml').text()) as Workflow;
  const scripts = await Promise.all(
    publishedFolders(workflow).map(async (folder) => {
      const manifest = (await Bun.file(`${folder}/package.json`).json()) as { readonly scripts?: Readonly<Record<string, string>> };
      return [folder, manifest.scripts?.['check:docs']] as const;
    })
  );
  const found = [...problems(workspaces, workflow), ...docsProblems(workflow, Object.fromEntries(scripts))];
  for (const problem of found) process.stderr.write(`check-ci-matrix: ${problem}\n`);
  if (found.length > 0) return 1;
  process.stdout.write(
    `check-ci-matrix: the ${JOB} matrix covers every workspace folder (${workspaces.join(', ')}); ${SHARDED} mutates in the shards of the ${MUTATION_JOB} job only; ${publishedFolders(workflow).join(', ')} check their command reference after the build.\n`
  );
  return 0;
};

process.exit(await main());
