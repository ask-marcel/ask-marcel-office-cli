#!/usr/bin/env bun
/*
 * The package boundaries on the import graph (package split, step 15). Write
 * will bundle core and its own modules, never read's; read will bundle core and
 * its own modules, never write's. So, with type imports and dynamic imports
 * counted (a type still needs its module in the package):
 *
 *   - a write module (a write command or a write-only helper) reaches only core
 *     and write modules;
 *   - a read command and a core module never reach a write module;
 *   - a core module reaches only core modules (phase 3: core builds alone);
 *   - no module outside auth reaches an auth module, by a relative path or by
 *     the package name (phase 4: read and write get tokens from the helper);
 *   - an auth module reaches only core and auth modules, and the package.json
 *     of its own package (phase 4: auth builds alone).
 *
 * Phase 3 moved core from src/ to core/src/, and phase 4 moves auth to
 * auth/src/. A file under core/ or auth/ gets its owner from its folder; the
 * lists name `src/...` and apply to the root src/ only (else SINGLE_PACKAGE_FILES
 * `src/{index,main}.ts` would match auth/src/main.ts). The root imports core by
 * its workspace name (`@ask-marcel/office-core/<path>`, read here as
 * core/src/<path>), and `@ask-marcel/office-auth/<path>` is read as
 * auth/src/<path>. Any other specifier of a workspace package is refused: the
 * hoisted workspace resolves it silently, and the check does not read that
 * package's folder.
 *
 * Ownership comes from eslint.config.js (WRITE_FILES, CORE_FILES,
 * CORE_READ_ONLY_FILES), the lists the core and read-only lint bans use, so a
 * file has one owner everywhere. Auth has no list: its folder decides. The read commands are the registry's commands
 * whose effect is `read`; every other registry command must be a write module.
 *
 * Test files (step 16): each belongs to one package, so a later phase moves it
 * whole. A test file imports no module of another package; core and the single
 * package's wiring (SINGLE_PACKAGE_FILES: its registry, shells and full client,
 * which phases 4 to 6 replace) are open to every test. Its package is the side
 * its name gives (`-write.test.ts`, `-auth.test.ts`), else the owner of the
 * module it is named after, else the one package its imports name, else read
 * when it imports the single package's wiring (the real-registry suites go with
 * read, step 13), else core. A module under auth/ is auth's; a module in no
 * list is read's. A test under core/ imports core modules only: core runs its
 * tests alone. A test under auth/ imports auth and core modules only, for the
 * same reason. A test outside core/ never falls to core (phase 3 moved every
 * core test into core/): one the rules give core is reported, so it moves to
 * core/src/ or is named in SINGLE_PACKAGE_TESTS. In the same way, a test
 * outside auth/ never falls to auth (phase 4 moved every auth test into auth/).
 * Only imports are read: a case that runs another package's command through the
 * registry by name is not seen, so such a case goes in a test file named for its
 * side. `--test-map` prints each test file's package.
 *
 * Fixtures: each run also checks the tree with one file's source replaced by a
 * violating import, and the ownership lists with one core file put in a second
 * list. Each must be reported, so a check that can no longer see an import or a
 * doubled file fails instead of passing (rule 15.10: a gate only ever seen
 * green is a hypothesis).
 *
 * Exit codes:
 *   0  no boundary is crossed, and every fixture is reported
 *   1  a boundary is crossed, a fixture is missed, an import does not resolve,
 *      or the ownership lists disagree with the registry or with each other
 */

import { existsSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import ts from 'typescript';
import { CORE_FILES, CORE_READ_ONLY_FILES, SINGLE_PACKAGE_FILES, WRITE_FILES } from '../eslint.config.js';

const ROOT = resolve(import.meta.dir, '..');
const COMMANDS = 'src/use-cases/commands';
// The workspace folders whose files the check reads; each package name reaches its folder's src/.
const FOLDERS = ['core', 'auth'] as const;
const WORKSPACE_PACKAGE = /^@ask-marcel\/office-(core|auth)\/(.+)$/;

// Where a module of the root src/ is now: in core/ or auth/ once its phase moved it there.
const located = (path: string): string => {
  const folder = FOLDERS.find((name) => existsSync(join(ROOT, name, path)));
  return folder === undefined ? path : `${folder}/${path}`;
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
// Auth modules that the fixtures name, under auth/src/ since phase 4 moved them.
const AUTH_LADDER = located('src/infra/auth.ts');
const TOKEN_ARGS = located(`${COMMANDS}/token-args.ts`);

// Tests of the single package as a whole: they read the whole root src/, every
// package's commands included, so no one package takes them. Phases 4 to 6
// split them with the commands they read. (Core's own half of the placeholder
// sweep is core/src/use-cases/commands/bin-placeholders.test.ts.)
const SINGLE_PACKAGE_TESTS: ReadonlySet<string> = new Set(['src/use-cases/commands/bin-placeholders.test.ts']);

type Sources = ReadonlyMap<string, string>;
type Graph = ReadonlyMap<string, ReadonlyArray<string>>;
type Owners = { readonly core: ReadonlySet<string>; readonly auth: ReadonlySet<string>; readonly write: ReadonlySet<string>; readonly reads: ReadonlyArray<string> };
type Fixture = { readonly name: string; readonly path: string; readonly source: string; readonly reported: string };
type Owner = 'core' | 'auth' | 'read' | 'write' | 'single';
type OwnerOf = (path: string) => Owner;

const FIXTURES: ReadonlyArray<Fixture> = [
  {
    name: 'a write command that imports a read helper',
    path: `${COMMANDS}/create-mail-draft.ts`,
    source: "import type { ResolvedLink } from './sharepoint-link-extractor.ts';\nexport type F = ResolvedLink;\n",
    reported: `${COMMANDS}/create-mail-draft.ts reaches ${COMMANDS}/sharepoint-link-extractor.ts`,
  },
  {
    name: 'a read command that imports a write-only helper',
    path: `${COMMANDS}/list-drives.ts`,
    source: "import { slimDraftResult } from './draft-response.ts';\nexport const f = (): unknown => slimDraftResult;\n",
    reported: `${COMMANDS}/list-drives.ts reaches ${COMMANDS}/draft-response.ts`,
  },
  {
    name: 'a core module that loads a write command with a dynamic import',
    path: CORE_FIXTURE,
    source: `export const f = async (): Promise<unknown> => import('${specifierTo(CORE_FIXTURE, `${COMMANDS}/create-reply-draft.ts`)}');\n`,
    reported: `${CORE_FIXTURE} reaches ${COMMANDS}/create-reply-draft.ts`,
  },
  {
    name: 'a core module that imports a read helper',
    path: CORE_READER_FIXTURE,
    source: `import { extractSharepointUrls } from '${specifierTo(CORE_READER_FIXTURE, `${COMMANDS}/sharepoint-link-extractor.ts`)}';\nexport const f = extractSharepointUrls;\n`,
    reported: `${CORE_READER_FIXTURE} reaches ${COMMANDS}/sharepoint-link-extractor.ts`,
  },
  {
    name: 'a read command that imports an auth module by the package name',
    path: `${COMMANDS}/list-drives.ts`,
    source: "import type { Result } from '@ask-marcel/office-auth/index.ts';\nexport type F = Result<1, 2>;\n",
    reported: `${COMMANDS}/list-drives.ts reaches auth/src/index.ts`,
  },
  {
    name: 'a read helper that imports an auth module by a relative path',
    path: `${COMMANDS}/sharepoint-link-extractor.ts`,
    source: `import { parseTokenArgs } from '${specifierTo(`${COMMANDS}/sharepoint-link-extractor.ts`, TOKEN_ARGS)}';\nexport const f = parseTokenArgs;\n`,
    reported: `${COMMANDS}/sharepoint-link-extractor.ts reaches ${TOKEN_ARGS}`,
  },
  {
    name: 'an auth module that imports a read helper',
    path: TOKEN_ARGS,
    source: `import { extractSharepointUrls } from '${specifierTo(TOKEN_ARGS, `${COMMANDS}/sharepoint-link-extractor.ts`)}';\nexport const f = extractSharepointUrls;\n`,
    reported: `${TOKEN_ARGS} reaches ${COMMANDS}/sharepoint-link-extractor.ts`,
  },
];

// Each must be reported by the test-file rule. The first two are new files, the
// third replaces a real test's source.
const TEST_FIXTURES: ReadonlyArray<Fixture> = [
  {
    name: 'a write-side test that imports a read helper',
    path: `${COMMANDS}/fixture-write.test.ts`,
    source: "import { extractSharepointUrls } from './sharepoint-link-extractor.ts';\nexport const f = extractSharepointUrls;\n",
    reported: `${COMMANDS}/fixture-write.test.ts (write) imports ${COMMANDS}/sharepoint-link-extractor.ts (read)`,
  },
  {
    name: 'a test that imports modules of two packages',
    path: `${COMMANDS}/fixture-mixed.test.ts`,
    source: `import type { AuthManager } from '${specifierTo(`${COMMANDS}/fixture-mixed.test.ts`, AUTH_LADDER)}';\nimport { extractSharepointUrls } from './sharepoint-link-extractor.ts';\nexport type F = AuthManager;\nexport const f = extractSharepointUrls;\n`,
    reported: `${COMMANDS}/fixture-mixed.test.ts imports modules of auth and read`,
  },
  {
    name: 'a test under core/ that imports the single package wiring',
    path: CORE_TEST_FIXTURE,
    source: `import { commands } from '${specifierTo(CORE_TEST_FIXTURE, `${COMMANDS}/index.ts`)}';\nexport const f = commands;\n`,
    reported: `${CORE_TEST_FIXTURE} (core) imports ${COMMANDS}/index.ts (single)`,
  },
  {
    name: 'a test under auth/ that imports the single package wiring',
    path: AUTH_TEST_FIXTURE,
    source: `import { commands } from '${specifierTo(AUTH_TEST_FIXTURE, `${COMMANDS}/index.ts`)}';\nexport const f = commands;\n`,
    reported: `${AUTH_TEST_FIXTURE} (auth) imports ${COMMANDS}/index.ts (single)`,
  },
  {
    name: 'a test outside core/ that imports core modules only',
    path: `${COMMANDS}/fixture-core-outside.test.ts`,
    source: "import { ok } from '@ask-marcel/office-core/domain/result.ts';\nexport const f = ok;\n",
    reported: `${COMMANDS}/fixture-core-outside.test.ts (core) sits outside core/`,
  },
  {
    name: 'a test outside auth/ that imports auth modules only',
    path: `${COMMANDS}/fixture-auth-outside.test.ts`,
    source: "import { resolveAuthPaths } from '@ask-marcel/office-auth/index.ts';\nexport const f = resolveAuthPaths;\n",
    reported: `${COMMANDS}/fixture-auth-outside.test.ts (auth) sits outside auth/`,
  },
  {
    name: 'a test named after a write command that imports an auth module',
    path: `${COMMANDS}/create-mail-draft.test.ts`,
    source: `import { parseTokenArgs } from '${specifierTo(`${COMMANDS}/create-mail-draft.test.ts`, TOKEN_ARGS)}';\nexport const f = parseTokenArgs;\n`,
    reported: `${COMMANDS}/create-mail-draft.test.ts (write) imports ${TOKEN_ARGS} (auth)`,
  },
];

// Must be reported by the rule that refuses a file in two ownership lists: a core file put in SINGLE_PACKAGE_FILES.
const OVERLAP_FIXTURE = { name: 'a file in two ownership lists', path: CORE_FIXTURE, reported: `${CORE_FIXTURE} (core, single)` } as const;

// Must be refused: the hoisted workspace would resolve it, and the check would not see its module.
const SPECIFIER_FIXTURE = {
  name: 'an import of a workspace package outside core and auth',
  path: `${COMMANDS}/list-drives.ts`,
  specifier: '@ask-marcel/office-read/index.ts',
} as const;

// The files under src/ of the root (no folder) or of a workspace folder, repo-relative.
const scan = (folder?: string): ReadonlyArray<string> =>
  folder === undefined ? [...new Bun.Glob('src/**/*.ts').scanSync(ROOT)] : [...new Bun.Glob('src/**/*.ts').scanSync(join(ROOT, folder))].map((path) => `${folder}/${path}`);

// The lists name `src/...` and apply to the root src/ only; every file under
// the folder of the list's package belongs to it too.
const glob = (patterns: ReadonlyArray<string>, folder?: string): ReadonlySet<string> =>
  new Set(
    [...patterns.flatMap((pattern) => [...new Bun.Glob(pattern).scanSync(ROOT)]), ...(folder === undefined ? [] : scan(folder))].filter((path) => !path.endsWith('.test.ts'))
  );

const readSources = async (tests: boolean): Promise<Sources> => {
  const sources = new Map<string, string>();
  for (const path of [...scan(), ...FOLDERS.flatMap((folder) => scan(folder))]) {
    const test = path.endsWith('.test.ts');
    if (test !== tests || (!test && /^((core|auth)\/)?src\/test-helpers\//.test(path))) continue;
    sources.set(path, await Bun.file(join(ROOT, path)).text());
  }
  return sources;
};

// The repo-relative file a relative specifier names (a module, or data such as
// package.json), or undefined for a package import. A relative specifier that
// names no file, or a workspace package outside core and auth, is an error.
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

// The package.json of a file's own package: its folder's, or the root's.
const packageJsonOf = (path: string): string => {
  const folder = FOLDERS.find((name) => path.startsWith(`${name}/`));
  return folder === undefined ? 'package.json' : `${folder}/package.json`;
};

const violations = (graph: Graph, owners: Owners): ReadonlyArray<string> => {
  const writeReach = [...owners.write].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || owners.write.has(path)));
  const readReach = [...owners.reads, ...owners.core].flatMap((start) => crossing(graph, start, (path) => !owners.write.has(path)));
  const coreReach = [...owners.core].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path)));
  const outsideAuth = [...graph.keys()].filter((start) => !owners.auth.has(start));
  const authReached = outsideAuth.flatMap((start) => crossing(graph, start, (path) => !owners.auth.has(path)));
  const authReach = [...owners.auth].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || owners.auth.has(path) || path === packageJsonOf(start)));
  return [...new Set([...writeReach, ...readReach, ...coreReach, ...authReached, ...authReach])];
};

type Lists = { readonly core: ReadonlySet<string>; readonly auth: ReadonlySet<string>; readonly write: ReadonlySet<string>; readonly single: ReadonlySet<string> };

const PACKAGES: ReadonlyArray<Owner> = ['auth', 'read', 'write'];

const ownerOf = (lists: Lists, path: string): Owner => {
  if (lists.core.has(path)) return 'core';
  if (lists.auth.has(path)) return 'auth';
  if (lists.write.has(path)) return 'write';
  if (lists.single.has(path)) return 'single';
  return 'read';
};

// The files that two ownership lists both claim.
const overlaps = (lists: Lists): ReadonlyArray<string> => {
  const claims = Object.entries(lists).flatMap(([name, paths]) => [...paths].map((path) => [path, name] as const));
  const first = new Map<string, string>();
  const twice: string[] = [];
  for (const [path, name] of claims) {
    const earlier = first.get(path);
    if (earlier !== undefined) twice.push(`${path} (${earlier}, ${name})`);
    first.set(path, name);
  }
  return twice;
};

// The packages that a test file's imports name; core and the single wiring are open to all.
const importedPackages = (imports: ReadonlyArray<string>, owner: OwnerOf): ReadonlySet<Owner> => new Set(imports.map(owner).filter((role) => PACKAGES.includes(role)));

// A test file's package, or undefined when its imports name two packages and nothing else decides.
const testPackage = (path: string, imports: ReadonlyArray<string>, owner: OwnerOf): Owner | undefined => {
  if (path.startsWith('core/')) return 'core';
  if (path.startsWith('auth/')) return 'auth';
  if (SINGLE_PACKAGE_TESTS.has(path)) return 'single';
  const side = /-(write|auth)\.test\.ts$/.exec(path)?.[1];
  if (side === 'write' || side === 'auth') return side;
  const namesake = path.replace(/\.test\.ts$/, '.ts');
  if (existsSync(join(ROOT, namesake)) && PACKAGES.includes(owner(namesake))) return owner(namesake);
  const named = [...importedPackages(imports, owner)];
  if (named.length > 1) return undefined;
  return named[0] ?? (imports.some((target) => owner(target) === 'single') ? 'read' : 'core');
};

const testViolations = (tests: Graph, owner: OwnerOf): ReadonlyArray<string> =>
  [...tests].flatMap(([path, imports]) => {
    const own = testPackage(path, imports, owner);
    if (own === undefined) return [`${path} imports modules of ${[...importedPackages(imports, owner)].sort().join(' and ')}`];
    if (own === 'core' && !path.startsWith('core/')) return [`${path} (core) sits outside core/: move it to core/src/, or name it in SINGLE_PACKAGE_TESTS`];
    if (own === 'auth' && !path.startsWith('auth/')) return [`${path} (auth) sits outside auth/: move it to auth/src/`];
    // A test under core/ runs with core alone, and one under auth/ with auth and
    // core: the single wiring is not open to them.
    const closed = (target: string): boolean => PACKAGES.includes(owner(target)) || path.startsWith('core/') || (path.startsWith('auth/') && owner(target) !== 'core');
    const foreign = imports.filter((target) => closed(target) && owner(target) !== own);
    return [...new Set(foreign)].map((target) => `${path} (${own}) imports ${target} (${owner(target)})`);
  });

// The test graph keeps only the imports of modules under src/ (not package.json, not a package).
const buildTestGraph = (sources: Sources): Graph => new Map([...buildGraph(sources)].map(([path, imports]) => [path, imports.filter((target) => /^((core|auth)\/)?src\//.test(target))]));

type CommandModule = { readonly meta?: { readonly effect?: unknown }; readonly execute?: unknown };

// The registry's commands, split by effect.
const registryCommands = async (): Promise<{ readonly reads: ReadonlyArray<string>; readonly writes: ReadonlyArray<string> }> => {
  const reads: string[] = [];
  const writes: string[] = [];
  for (const file of new Bun.Glob('*.ts').scanSync(join(ROOT, COMMANDS))) {
    if (file.endsWith('.test.ts')) continue;
    const module = (await import(join(ROOT, COMMANDS, file))) as CommandModule;
    if (typeof module.execute !== 'function' || module.meta?.effect === undefined) continue;
    (module.meta.effect === 'read' ? reads : writes).push(`${COMMANDS}/${file}`);
  }
  return { reads, writes };
};

const main = async (): Promise<number> => {
  const core = glob([...CORE_FILES, ...CORE_READ_ONLY_FILES], 'core');
  const write = glob(WRITE_FILES);
  const lists: Lists = { core, auth: glob([], 'auth'), write, single: glob(SINGLE_PACKAGE_FILES) };
  const owner: OwnerOf = (path) => ownerOf(lists, path);
  const tests = await readSources(true);
  if (process.argv.includes('--test-map')) {
    console.log([...buildTestGraph(tests)].map(([path, imports]) => `${testPackage(path, imports, owner) ?? 'mixed'}\t${path}`).sort().join('\n'));
    return 0;
  }
  const { reads, writes } = await registryCommands();
  const unlisted = writes.filter((path) => !write.has(path));
  const misplaced = [...write].filter((path) => reads.includes(path) || core.has(path));
  const doubled = overlaps(lists);
  if (unlisted.length > 0 || misplaced.length > 0 || doubled.length > 0 || reads.length === 0) {
    console.error(`check-package-imports: the ownership lists disagree with the registry or with each other. Write commands outside WRITE_FILES: ${unlisted.join(', ') || 'none'}. WRITE_FILES that are read commands or core: ${misplaced.join(', ') || 'none'}. Files in two lists: ${doubled.join(', ') || 'none'}. Read commands found: ${reads.length}.`);
    return 1;
  }
  const owners: Owners = { core, auth: lists.auth, write, reads };
  const sources = await readSources(false);
  const found = [...violations(buildGraph(sources), owners), ...testViolations(buildTestGraph(tests), owner)];
  const doubledFixture = overlaps({ ...lists, single: new Set([...lists.single, OVERLAP_FIXTURE.path]) });
  const missed = [
    ...FIXTURES.filter((fixture) => !violations(buildGraph(new Map([...sources, [fixture.path, fixture.source]])), owners).some((line) => line.startsWith(fixture.reported))),
    ...TEST_FIXTURES.filter((fixture) => !testViolations(buildTestGraph(new Map([...tests, [fixture.path, fixture.source]])), owner).some((line) => line.startsWith(fixture.reported))),
    ...(doubledFixture.includes(OVERLAP_FIXTURE.reported) ? [] : [OVERLAP_FIXTURE]),
    ...(refuses(SPECIFIER_FIXTURE.path, SPECIFIER_FIXTURE.specifier) ? [] : [SPECIFIER_FIXTURE]),
  ];
  console.log(`check-package-imports: ${write.size} write modules, ${core.size} core modules, ${lists.auth.size} auth modules, ${reads.length} read commands, ${sources.size} modules and ${tests.size} test files read.`);
  if (found.length > 0) console.error(`check-package-imports: a package boundary is crossed:\n  ${found.join('\n  ')}`);
  if (missed.length > 0) console.error(`check-package-imports: the check missed its fixture(s): ${missed.map((fixture) => fixture.name).join('; ')}.`);
  if (found.length > 0 || missed.length > 0) return 1;
  console.log('check-package-imports: write reaches only core and write; no read command and no core module reaches write; core reaches only core; only auth reaches auth, and auth reaches only core and auth; no test file imports a module of another package; every core test sits in core/ and every auth test in auth/; every fixture is reported.');
  return 0;
};

process.exit(await main());
