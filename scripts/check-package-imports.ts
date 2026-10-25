#!/usr/bin/env bun
/*
 * The package boundaries on the import graph (package split, step 15). Write
 * will bundle core and its own modules, never read's; read will bundle core and
 * its own modules, never write's. So, with type imports and dynamic imports
 * counted (a type still needs its module in the package):
 *
 *   - a write module (a write command, a write-only helper or the write
 *     package's wiring) reaches only core and write modules, and the
 *     package.json of its own package (the write bin reads its version there);
 *   - no module outside write reaches a write module, by a relative path or by
 *     the package name (phase 5: read and auth build without write);
 *   - a core module reaches only core modules (phase 3: core builds alone);
 *   - no module outside auth reaches an auth module, by a relative path or by
 *     the package name (phase 4: read and write get tokens from the helper);
 *   - an auth module reaches only core and auth modules, and the package.json
 *     of its own package (phase 4: auth builds alone);
 *   - a read module reaches only core and read modules, and the package.json
 *     of its own package (phase 6: read builds alone).
 *
 * Phases 3 to 6 moved core, auth, write and read into core/src/, auth/src/,
 * write/src/ and read/src/, and the root holds no source. The folder decides
 * the package of every file (LESSONS 2026-10-18): a file under core/, auth/,
 * read/ or write/ belongs to that folder's package, and no list names a file.
 * A package imports core by its workspace name
 * (`@ask-marcel/office-core/<path>`, read here as core/src/<path>), and
 * `@ask-marcel/office-auth/<path>`, `@ask-marcel/office-read/<path>` and
 * `@ask-marcel/office-write/<path>` are read as auth/src/<path>,
 * read/src/<path> and write/src/<path>. Any other specifier of a workspace
 * package is refused: the hoisted workspace resolves it silently, and the check
 * does not read that package's folder. The read commands are the commands of
 * read's and write's command folders whose effect is `read`: each must sit
 * under read/, and every other command under write/.
 *
 * Test files (step 16): each belongs to the package of its folder, and a test
 * outside every workspace folder is reported. A test under core/ imports core
 * modules only: core runs its tests alone. A test under auth/, read/ or write/
 * imports its own package and core only, for the same reason. Only imports are
 * read: a case that runs another package's command through a registry by name
 * is not seen. `--test-map` prints each test file's package.
 *
 * Fixtures: each run also checks the tree with one file's source replaced by a
 * violating import. Each must be reported, so a check that can no longer see an
 * import fails instead of passing (rule 15.10: a gate only ever seen
 * green is a hypothesis).
 *
 * Exit codes:
 *   0  no boundary is crossed, and every fixture is reported
 *   1  a boundary is crossed, a fixture is missed, an import does not resolve,
 *      or a command sits in the folder of the wrong package
 */

import { existsSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import ts from 'typescript';

const ROOT = resolve(import.meta.dir, '..');
const COMMANDS = 'src/use-cases/commands';
// The workspace folders whose files the check reads; each package name reaches its folder's src/.
const FOLDERS = ['core', 'auth', 'read', 'write'] as const;
const WORKSPACE_PACKAGE = /^@ask-marcel\/office-(core|auth|read|write)\/(.+)$/;

type Folder = (typeof FOLDERS)[number];

// The workspace folder that holds a fixture's module. A path that no folder holds is an error, so a
// fixture never names a file that is gone.
const located = (path: string): string => {
  const folder = FOLDERS.find((name) => existsSync(join(ROOT, name, path)));
  if (folder === undefined) throw new Error(`no workspace folder holds ${path}`);
  return `${folder}/${path}`;
};
// A relative specifier from `from` to the module `to`, both repo-relative.
const specifierTo = (from: string, to: string): string => {
  const path = relative(dirname(from), to);
  return path.startsWith('.') ? path : `./${path}`;
};
const CORE_FIXTURE = located(`${COMMANDS}/format-zod-error.ts`);
const CORE_READER_FIXTURE = located(`${COMMANDS}/quote-boundary.ts`);
const CORE_TEST_FIXTURE = 'core/src/use-cases/commands/fixture-core.test.ts';
const AUTH_TEST_FIXTURE = 'auth/src/composition/fixture-auth.test.ts';
const WRITE_FOLDER_TEST_FIXTURE = 'write/src/composition/fixture-write-folder.test.ts';
// Auth modules that the fixtures name, under auth/src/ since phase 4 moved them.
const TOKEN_ARGS = located(`${COMMANDS}/token-args.ts`);
// Write modules and tests that the fixtures name, under write/src/ since phase 5 moved them.
const MAIL_DRAFT = located(`${COMMANDS}/create-mail-draft.ts`);
const MAIL_DRAFT_TEST = located(`${COMMANDS}/create-mail-draft.test.ts`);
const DRAFT_RESPONSE = located(`${COMMANDS}/draft-response.ts`);
const REPLY_DRAFT = located(`${COMMANDS}/create-reply-draft.ts`);
const WRITE_TEST_FIXTURE = `${dirname(MAIL_DRAFT)}/fixture-write.test.ts`;
// Read modules and tests that the fixtures name, under read/src/ since phase 6 moved them.
const LIST_DRIVES = located(`${COMMANDS}/list-drives.ts`);
const LINK_EXTRACTOR = located(`${COMMANDS}/sharepoint-link-extractor.ts`);
const REGISTRY = located(`${COMMANDS}/index.ts`);
const READ_TEST_FIXTURE = 'read/src/composition/fixture-read.test.ts';
// A new test file in no workspace folder.
const OUTSIDE_TEST_FIXTURE = `${COMMANDS}/fixture-outside.test.ts`;

type Sources = ReadonlyMap<string, string>;
type Graph = ReadonlyMap<string, ReadonlyArray<string>>;
type Owners = Readonly<Record<Folder, ReadonlySet<string>>>;
type Fixture = { readonly name: string; readonly path: string; readonly source: string; readonly reported: string };

const FIXTURES: ReadonlyArray<Fixture> = [
  {
    name: 'a write command that imports a read helper',
    path: MAIL_DRAFT,
    source: `import type { ResolvedLink } from '${specifierTo(MAIL_DRAFT, LINK_EXTRACTOR)}';\nexport type F = ResolvedLink;\n`,
    reported: `${MAIL_DRAFT} reaches ${LINK_EXTRACTOR}`,
  },
  {
    name: 'a read command that imports a write-only helper',
    path: LIST_DRIVES,
    source: `import { slimDraftResult } from '${specifierTo(LIST_DRIVES, DRAFT_RESPONSE)}';\nexport const f = (): unknown => slimDraftResult;\n`,
    reported: `${LIST_DRIVES} reaches ${DRAFT_RESPONSE}`,
  },
  {
    name: 'a core module that loads a write command with a dynamic import',
    path: CORE_FIXTURE,
    source: `export const f = async (): Promise<unknown> => import('${specifierTo(CORE_FIXTURE, REPLY_DRAFT)}');\n`,
    reported: `${CORE_FIXTURE} reaches ${REPLY_DRAFT}`,
  },
  {
    name: 'a core module that imports a read helper',
    path: CORE_READER_FIXTURE,
    source: `import { extractSharepointUrls } from '${specifierTo(CORE_READER_FIXTURE, LINK_EXTRACTOR)}';\nexport const f = extractSharepointUrls;\n`,
    reported: `${CORE_READER_FIXTURE} reaches ${LINK_EXTRACTOR}`,
  },
  {
    name: 'a read command that imports an auth module by the package name',
    path: LIST_DRIVES,
    source: "import type { Result } from '@ask-marcel/office-auth/index.ts';\nexport type F = Result<1, 2>;\n",
    reported: `${LIST_DRIVES} reaches auth/src/index.ts`,
  },
  {
    name: 'a read helper that imports an auth module by a relative path',
    path: LINK_EXTRACTOR,
    source: `import { parseTokenArgs } from '${specifierTo(LINK_EXTRACTOR, TOKEN_ARGS)}';\nexport const f = parseTokenArgs;\n`,
    reported: `${LINK_EXTRACTOR} reaches ${TOKEN_ARGS}`,
  },
  {
    name: 'an auth module that imports a read helper',
    path: TOKEN_ARGS,
    source: `import { extractSharepointUrls } from '${specifierTo(TOKEN_ARGS, LINK_EXTRACTOR)}';\nexport const f = extractSharepointUrls;\n`,
    reported: `${TOKEN_ARGS} reaches ${LINK_EXTRACTOR}`,
  },
  {
    name: 'a read command that imports a write module by the package name',
    path: LIST_DRIVES,
    source: "import type { Result } from '@ask-marcel/office-write/index.ts';\nexport type F = Result<1, 2>;\n",
    reported: `${LIST_DRIVES} reaches write/src/index.ts`,
  },
  {
    name: 'the read registry that imports a write command',
    path: REGISTRY,
    source: `import * as draft from '${specifierTo(REGISTRY, MAIL_DRAFT)}';\nexport const f = draft;\n`,
    reported: `${REGISTRY} reaches ${MAIL_DRAFT}`,
  },
  {
    name: 'a read module that imports the package.json of another package',
    path: LIST_DRIVES,
    source: `import pkg from '${specifierTo(LIST_DRIVES, 'write/package.json')}' with { type: 'json' };\nexport const f = pkg;\n`,
    reported: `${LIST_DRIVES} reaches write/package.json`,
  },
  {
    name: 'a write module that imports an auth module',
    path: DRAFT_RESPONSE,
    source: `import { parseTokenArgs } from '${specifierTo(DRAFT_RESPONSE, TOKEN_ARGS)}';\nexport const f = parseTokenArgs;\n`,
    reported: `${DRAFT_RESPONSE} reaches ${TOKEN_ARGS}`,
  },
  {
    name: 'a write module that imports the package.json of another package',
    path: DRAFT_RESPONSE,
    source: `import pkg from '${specifierTo(DRAFT_RESPONSE, 'core/package.json')}' with { type: 'json' };\nexport const f = pkg;\n`,
    reported: `${DRAFT_RESPONSE} reaches core/package.json`,
  },
  {
    name: 'an auth module that imports a write module',
    path: TOKEN_ARGS,
    source: `import { slimDraftResult } from '${specifierTo(TOKEN_ARGS, DRAFT_RESPONSE)}';\nexport const f = slimDraftResult;\n`,
    reported: `${TOKEN_ARGS} reaches ${DRAFT_RESPONSE}`,
  },
];

// Each must be reported by the test-file rule. The last replaces a real test's
// source; the others are new files.
const TEST_FIXTURES: ReadonlyArray<Fixture> = [
  {
    name: 'a write-side test that imports a read helper',
    path: WRITE_TEST_FIXTURE,
    source: `import { extractSharepointUrls } from '${specifierTo(WRITE_TEST_FIXTURE, LINK_EXTRACTOR)}';\nexport const f = extractSharepointUrls;\n`,
    reported: `${WRITE_TEST_FIXTURE} (write) imports ${LINK_EXTRACTOR} (read)`,
  },
  {
    name: 'a test under read/ that imports a write module',
    path: READ_TEST_FIXTURE,
    source: `import { slimDraftResult } from '${specifierTo(READ_TEST_FIXTURE, DRAFT_RESPONSE)}';\nexport const f = slimDraftResult;\n`,
    reported: `${READ_TEST_FIXTURE} (read) imports ${DRAFT_RESPONSE} (write)`,
  },
  {
    name: 'a test under core/ that imports the read registry',
    path: CORE_TEST_FIXTURE,
    source: `import { commands } from '${specifierTo(CORE_TEST_FIXTURE, REGISTRY)}';\nexport const f = commands;\n`,
    reported: `${CORE_TEST_FIXTURE} (core) imports ${REGISTRY} (read)`,
  },
  {
    name: 'a test under auth/ that imports the read registry',
    path: AUTH_TEST_FIXTURE,
    source: `import { commands } from '${specifierTo(AUTH_TEST_FIXTURE, REGISTRY)}';\nexport const f = commands;\n`,
    reported: `${AUTH_TEST_FIXTURE} (auth) imports ${REGISTRY} (read)`,
  },
  {
    name: 'a test under write/ that imports a read helper',
    path: WRITE_FOLDER_TEST_FIXTURE,
    source: `import { extractSharepointUrls } from '${specifierTo(WRITE_FOLDER_TEST_FIXTURE, LINK_EXTRACTOR)}';\nexport const f = extractSharepointUrls;\n`,
    reported: `${WRITE_FOLDER_TEST_FIXTURE} (write) imports ${LINK_EXTRACTOR} (read)`,
  },
  {
    name: 'a test under write/ that imports the read registry',
    path: WRITE_FOLDER_TEST_FIXTURE,
    source: `import { commands } from '${specifierTo(WRITE_FOLDER_TEST_FIXTURE, REGISTRY)}';\nexport const f = commands;\n`,
    reported: `${WRITE_FOLDER_TEST_FIXTURE} (write) imports ${REGISTRY} (read)`,
  },
  {
    name: 'a test outside every workspace folder',
    path: OUTSIDE_TEST_FIXTURE,
    source: "import { ok } from '@ask-marcel/office-core/domain/result.ts';\nexport const f = ok;\n",
    reported: `${OUTSIDE_TEST_FIXTURE} sits outside every workspace folder`,
  },
  {
    name: 'a test named after a write command that imports an auth module',
    path: MAIL_DRAFT_TEST,
    source: `import { parseTokenArgs } from '${specifierTo(MAIL_DRAFT_TEST, TOKEN_ARGS)}';\nexport const f = parseTokenArgs;\n`,
    reported: `${MAIL_DRAFT_TEST} (write) imports ${TOKEN_ARGS} (auth)`,
  },
];

// Must be refused: the hoisted workspace would resolve it, and the check would not see its module.
const SPECIFIER_FIXTURE = {
  name: 'an import of a workspace package outside core, auth, read and write',
  path: LIST_DRIVES,
  specifier: '@ask-marcel/office-other/index.ts',
} as const;

// The files under src/ of a workspace folder, repo-relative.
const scan = (folder: Folder): ReadonlyArray<string> => [...new Bun.Glob('src/**/*.ts').scanSync(join(ROOT, folder))].map((path) => `${folder}/${path}`);

// The modules of a package: every file under its folder's src/ but the tests.
const modulesOf = (folder: Folder): ReadonlySet<string> => new Set(scan(folder).filter((path) => !path.endsWith('.test.ts')));

// The workspace folder of a repo-relative path, or undefined for a path outside every folder.
const folderOf = (path: string): Folder | undefined => FOLDERS.find((name) => path.startsWith(`${name}/`));

const readSources = async (tests: boolean): Promise<Sources> => {
  const sources = new Map<string, string>();
  for (const path of FOLDERS.flatMap((folder) => scan(folder))) {
    const test = path.endsWith('.test.ts');
    if (test !== tests || (!test && /^(core|auth|read|write)\/src\/test-helpers\//.test(path))) continue;
    sources.set(path, await Bun.file(join(ROOT, path)).text());
  }
  return sources;
};

// The repo-relative file a relative specifier names (a module, or data such as
// package.json), or undefined for a package import. A relative specifier that
// names no file, or a workspace package outside core, auth, read and write, is an error.
const resolveSpecifier = (from: string, specifier: string): string | undefined => {
  const workspace = WORKSPACE_PACKAGE.exec(specifier);
  if (workspace !== null) return resolveSpecifier(`${workspace[1]}/src/index.ts`, `./${workspace[2]}`);
  if (specifier.startsWith('@ask-marcel/')) throw new Error(`${from}: the import '${specifier}' names a workspace package whose folder this check does not read`);
  if (!specifier.startsWith('.')) return undefined;
  const path = relative(ROOT, resolve(ROOT, dirname(from), specifier));
  const candidates = [path, `${path}.ts`, join(path, 'index.ts')];
  const found = candidates.find((candidate) => /\.[a-z]+$/.test(candidate) && existsSync(join(ROOT, candidate)));
  if (found === undefined) throw new Error(`${from}: the import '${specifier}' names no module`);
  return found;
};

const refuses = (from: string, specifier: string): boolean => {
  try {
    resolveSpecifier(from, specifier);
    return false;
  } catch {
    return true;
  }
};

const buildGraph = (sources: Sources): Graph =>
  new Map(
    [...sources].map(([path, source]) => [
      path,
      ts
        .preProcessFile(source, true, true)
        .importedFiles.map((file) => resolveSpecifier(path, file.fileName))
        .filter((target): target is string => target !== undefined),
    ])
  );

// Every module `start` reaches, itself excluded, with the first importer seen.
const reach = (graph: Graph, start: string): ReadonlyMap<string, string> => {
  const via = new Map<string, string>();
  const queue = [start];
  for (let at = queue.shift(); at !== undefined; at = queue.shift()) {
    for (const target of graph.get(at) ?? []) {
      if (target === start || via.has(target)) continue;
      via.set(target, at);
      queue.push(target);
    }
  }
  return via;
};

const crossing = (graph: Graph, start: string, allowed: (path: string) => boolean): ReadonlyArray<string> =>
  [...reach(graph, start)].filter(([target]) => !allowed(target)).map(([target, importer]) => `${start} reaches ${target} (imported by ${importer})`);

// The package.json of a file's own package: its folder's.
const packageJsonOf = (path: string): string => `${folderOf(path) ?? '(no folder)'}/package.json`;

const violations = (graph: Graph, owners: Owners): ReadonlyArray<string> => {
  const writeReach = [...owners.write].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || owners.write.has(path) || path === packageJsonOf(start)));
  const outsideWrite = [...graph.keys()].filter((start) => !owners.write.has(start));
  const writeReached = outsideWrite.flatMap((start) => crossing(graph, start, (path) => !owners.write.has(path)));
  const coreReach = [...owners.core].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path)));
  const outsideAuth = [...graph.keys()].filter((start) => !owners.auth.has(start));
  const authReached = outsideAuth.flatMap((start) => crossing(graph, start, (path) => !owners.auth.has(path)));
  const authReach = [...owners.auth].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || owners.auth.has(path) || path === packageJsonOf(start)));
  const readReach = [...owners.read].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || owners.read.has(path) || path === packageJsonOf(start)));
  return [...new Set([...writeReach, ...writeReached, ...coreReach, ...authReached, ...authReach, ...readReach])];
};

// A test runs with its own package and core only (core's with core only), so it imports no other package.
const testViolations = (tests: Graph): ReadonlyArray<string> =>
  [...tests].flatMap(([path, imports]) => {
    const own = folderOf(path);
    if (own === undefined) return [`${path} sits outside every workspace folder: move it to the src/ of its package`];
    const foreign = imports.filter((target) => folderOf(target) !== own && folderOf(target) !== 'core');
    return [...new Set(foreign)].map((target) => `${path} (${own}) imports ${target} (${folderOf(target) ?? 'no folder'})`);
  });

// The test graph keeps only the imports of modules under a folder's src/ (not package.json, not a package).
const buildTestGraph = (sources: Sources): Graph =>
  new Map([...buildGraph(sources)].map(([path, imports]) => [path, imports.filter((target) => /^(core|auth|read|write)\/src\//.test(target))]));

type CommandModule = { readonly meta?: { readonly effect?: unknown }; readonly execute?: unknown };

// The command folders of read and write.
const COMMAND_FOLDERS = [`read/${COMMANDS}`, `write/${COMMANDS}`];

// The commands of the read registry and of the write package, split by effect.
const registryCommands = async (): Promise<{ readonly reads: ReadonlyArray<string>; readonly writes: ReadonlyArray<string> }> => {
  const reads: string[] = [];
  const writes: string[] = [];
  for (const folder of COMMAND_FOLDERS) {
    for (const file of new Bun.Glob('*.ts').scanSync(join(ROOT, folder))) {
      if (file.endsWith('.test.ts')) continue;
      const module = (await import(join(ROOT, folder, file))) as CommandModule;
      if (typeof module.execute !== 'function' || module.meta?.effect === undefined) continue;
      (module.meta.effect === 'read' ? reads : writes).push(`${folder}/${file}`);
    }
  }
  return { reads, writes };
};

const main = async (): Promise<number> => {
  const owners: Owners = { core: modulesOf('core'), auth: modulesOf('auth'), read: modulesOf('read'), write: modulesOf('write') };
  const tests = await readSources(true);
  if (process.argv.includes('--test-map')) {
    console.log(
      [...tests.keys()]
        .map((path) => `${folderOf(path) ?? 'no folder'}\t${path}`)
        .sort()
        .join('\n')
    );
    return 0;
  }
  const { reads, writes } = await registryCommands();
  const misplaced = [...reads.filter((path) => !owners.read.has(path)), ...writes.filter((path) => !owners.write.has(path))];
  if (misplaced.length > 0 || reads.length === 0) {
    console.error(`check-package-imports: a command sits in the folder of the wrong package: ${misplaced.join(', ') || 'none'}. Read commands found: ${reads.length}.`);
    return 1;
  }
  const sources = await readSources(false);
  const found = [...violations(buildGraph(sources), owners), ...testViolations(buildTestGraph(tests))];
  const missed = [
    ...FIXTURES.filter((fixture) => !violations(buildGraph(new Map([...sources, [fixture.path, fixture.source]])), owners).some((line) => line.startsWith(fixture.reported))),
    ...TEST_FIXTURES.filter((fixture) => !testViolations(buildTestGraph(new Map([...tests, [fixture.path, fixture.source]]))).some((line) => line.startsWith(fixture.reported))),
    ...(refuses(SPECIFIER_FIXTURE.path, SPECIFIER_FIXTURE.specifier) ? [] : [SPECIFIER_FIXTURE]),
  ];
  console.log(
    `check-package-imports: ${owners.write.size} write modules, ${owners.core.size} core modules, ${owners.auth.size} auth modules, ${owners.read.size} read modules, ${reads.length} read commands, ${sources.size} modules and ${tests.size} test files read.`
  );
  if (found.length > 0) console.error(`check-package-imports: a package boundary is crossed:\n  ${found.join('\n  ')}`);
  if (missed.length > 0) console.error(`check-package-imports: the check missed its fixture(s): ${missed.map((fixture) => fixture.name).join('; ')}.`);
  if (found.length > 0 || missed.length > 0) return 1;
  console.log(
    'check-package-imports: only write reaches write, and write reaches only core and write; core reaches only core; only auth reaches auth, and auth reaches only core and auth; read reaches only core and read; no test file imports a module of another package, and every test sits in the folder of its package; every fixture is reported.'
  );
  return 0;
};

process.exit(await main());
