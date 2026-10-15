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
 *   --selftest  proves a missing folder, an extra folder and a missing matrix
 *               are rejected, and an equal one passes (rule 15.10)
 *
 * Exit codes: 0 the matrix lists exactly the workspaces; 1 it does not.
 */

type Workflow = {
  readonly jobs?: Readonly<Record<string, { readonly strategy?: { readonly matrix?: { readonly include?: ReadonlyArray<{ readonly folder?: string }> } } }>>;
};

const JOB = 'workspace';

const matrixFolders = (workflow: Workflow): ReadonlyArray<string> | undefined => {
  const include = workflow.jobs?.[JOB]?.strategy?.matrix?.include;
  return include?.map((entry) => entry.folder ?? '');
};

const problems = (workspaces: ReadonlyArray<string>, workflow: Workflow): ReadonlyArray<string> => {
  const folders = matrixFolders(workflow);
  if (folders === undefined) return [`ci.yml has no jobs.${JOB}.strategy.matrix.include`];
  return [
    ...workspaces.filter((w) => !folders.includes(w)).map((w) => `workspace folder ${w} has no job in the ${JOB} matrix`),
    ...folders.filter((f) => !workspaces.includes(f)).map((f) => `the ${JOB} matrix names ${f || '(no folder)'}, which is not a workspace`),
  ];
};

const selftest = (): number => {
  const matrix = (folders: ReadonlyArray<string>): Workflow => ({ jobs: { [JOB]: { strategy: { matrix: { include: folders.map((folder) => ({ folder })) } } } } });
  const cases = [
    { name: 'an equal matrix', found: problems(['core', 'auth'], matrix(['auth', 'core'])).length, expected: 0 },
    { name: 'a missing folder', found: problems(['core', 'auth', 'read'], matrix(['core', 'auth'])).length, expected: 1 },
    { name: 'an extra folder', found: problems(['core'], matrix(['core', 'read'])).length, expected: 1 },
    { name: 'no matrix', found: problems(['core'], { jobs: {} }).length, expected: 1 },
  ];
  const wrong = cases.filter((c) => c.found !== c.expected);
  for (const c of wrong) process.stderr.write(`check-ci-matrix: SELFTEST FAILED, ${c.name} gave ${c.found} problem(s), expected ${c.expected}\n`);
  if (wrong.length === 0) process.stdout.write('check-ci-matrix: selftest passed (a missing folder, an extra folder and no matrix are rejected).\n');
  return wrong.length === 0 ? 0 : 1;
};

const main = async (): Promise<number> => {
  if (process.argv.includes('--selftest')) return selftest();
  const { workspaces } = (await Bun.file('package.json').json()) as { readonly workspaces: ReadonlyArray<string> };
  const workflow = Bun.YAML.parse(await Bun.file('.github/workflows/ci.yml').text()) as Workflow;
  const found = problems(workspaces, workflow);
  for (const problem of found) process.stderr.write(`check-ci-matrix: ${problem}\n`);
  if (found.length > 0) return 1;
  process.stdout.write(`check-ci-matrix: the ${JOB} matrix covers every workspace folder (${workspaces.join(', ')}).\n`);
  return 0;
};

process.exit(await main());
