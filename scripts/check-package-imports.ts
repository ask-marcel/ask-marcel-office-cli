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
 *     the package name (phase 5: the root and auth build without write);
 *   - a core module reaches only core modules (phase 3: core builds alone);
 *   - no module outside auth reaches an auth module, by a relative path or by
 *     the package name (phase 4: read and write get tokens from the helper);
 *   - an auth module reaches only core and auth modules, and the package.json
 *     of its own package (phase 4: auth builds alone);
 *   - a read module reaches only core and read modules, and the package.json
 *     of its own package: the root's while read is in the root src/, then
 *     read/package.json (phase 6: read builds alone).
 *
 * Phase 3 moved core from src/ to core/src/, phase 4 moved auth to auth/src/,
 * phase 5 moved write to write/src/, and phase 6 moves read to read/src/. A
 * file under core/, auth/, read/ or write/ gets its owner from its folder; the
 * lists name `src/...` and apply to the root src/ only (else
 * SINGLE_PACKAGE_FILES `src/{index,main}.ts` would match auth/src/main.ts,
 * read/src/main.ts and write/src/main.ts). The root imports core by its
 * workspace name (`@ask-marcel/office-core/<path>`, read here as
 * core/src/<path>), and `@ask-marcel/office-auth/<path>`,
 * `@ask-marcel/office-read/<path>` and `@ask-marcel/office-write/<path>` are
 * read as auth/src/<path>, read/src/<path> and write/src/<path>. Any other
 * specifier of a workspace package is refused: the hoisted workspace resolves
 * it silently, and the check does not read that package's folder. The fixtures
 * find each read module that they name in the root src/ or in read/src/,
 * wherever it is.
 *
 * Ownership comes from eslint.config.js (CORE_FILES, CORE_READ_ONLY_FILES),
 * the lists the core and read-only lint bans use, so a file has one owner
 * everywhere. Auth and write have no list: their folders decide (LESSONS
 * 2026-10-18). The read commands are the commands (of the root registry folder
 * and of write's) whose effect is `read`; every other command must be a write
 * module, so it must sit under write/.
 *
 * Test files (step 16): each belongs to one package, so a later phase moves it
 * whole. A test file imports no module of another package; core and the single
 * package's wiring (SINGLE_PACKAGE_FILES: its registry, shells and full client,
 * which phases 4 to 6 replace) are open to every test. Its package is the side
 * its name gives (`-write.test.ts`, `-auth.test.ts`), else the owner of the
 * module it is named after, else the one package its imports name, else read
 * when it imports the single package's wiring (the real-registry suites go with
 * read, step 13), else core. A module under auth/ is auth's, a module or a test
 * under write/ is write's, and a module or a test under read/ is read's; a
 * module in no list is read's. A test under
 * core/ imports core modules only: core runs its tests alone. A test under
 * auth/ imports auth and core modules only, and a test under write/ imports
 * write and core modules only, for the same reason. A test outside core/ never falls to core (phase 3 moved every
 * core test into core/): one the rules give core is reported, so it moves to
 * core/src/ or is named in SINGLE_PACKAGE_TESTS. In the same way, a test
 * outside auth/ never falls to auth (phase 4 moved every auth test into auth/),
 * and a test outside write/ never falls to write (phase 5 moved every write test
 * into write/).
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
import { CORE_FILES, CORE_READ_ONLY_FILES, SINGLE_PACKAGE_FILES } from '../eslint.config.js';

const ROOT = resolve(import.meta.dir, '..');
const COMMANDS = 'src/use-cases/commands';
// The workspace folders whose files the check reads; each package name reaches its folder's src/.
const FOLDERS = ['core', 'auth', 'read', 'write'] as const;
const WORKSPACE_PACKAGE = /^@ask-marcel\/office-(core|auth|read|write)\/(.+)$/;

// Where a module of the root src/ is now: in core/, auth/, read/ or write/ once its phase moved it there.
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
const WRITE_FOLDER_TEST_FIXTURE = 'write/src/composition/fixture-write-folder.test.ts';
// Auth modules that the fixtures name, under auth/src/ since phase 4 moved them.
const AUTH_LADDER = located('src/infra/auth.ts');
const TOKEN_ARGS = located(`${COMMANDS}/token-args.ts`);
// Write modules and tests that the fixtures name, under write/src/ since phase 5 moved them.
const MAIL_DRAFT = located(`${COMMANDS}/create-mail-draft.ts`);
const MAIL_DRAFT_TEST = located(`${COMMANDS}/create-mail-draft.test.ts`);
const DRAFT_RESPONSE = located(`${COMMANDS}/draft-response.ts`);
const REPLY_DRAFT = located(`${COMMANDS}/create-reply-draft.ts`);
const WRITE_TEST_FIXTURE = `${dirname(MAIL_DRAFT)}/fixture-write.test.ts`;
// Read modules that the fixtures name: in the root src/ until phase 6 moves them to read/src/.
const LIST_DRIVES = located(`${COMMANDS}/list-drives.ts`);
const LINK_EXTRACTOR = located(`${COMMANDS}/sharepoint-link-extractor.ts`);
const REGISTRY = located(`${COMMANDS}/index.ts`);
// New test files with no folder: the rules give them their package from their imports alone.
const MIXED_TEST_FIXTURE = `${COMMANDS}/fixture-mixed.test.ts`;

// Tests of the single package as a whole: they read the whole root src/, every
// package's commands included, so no one package takes them. Phases 4 to 6
// split them with the commands they read. (Core's own half of the placeholder
// sweep is core/src/use-cases/commands/bin-placeholders.test.ts.)
const SINGLE_PACKAGE_TESTS: ReadonlySet<string> = new Set(['src/use-cases/commands/bin-placeholders.test.ts']);

type Sources = ReadonlyMap<string, string>;
type Graph = ReadonlyMap<string, ReadonlyArray<string>>;
type Owners = { readonly core: ReadonlySet<string>; readonly auth: ReadonlySet<string>; readonly write: ReadonlySet<string> };
type Fixture = { readonly name: string; readonly path: string; readonly source: string; readonly reported: string };
type Owner = 'core' | 'auth' | 'read' | 'write' | 'single';
type OwnerOf = (path: string) => Owner;

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

// Each must be reported by the test-file rule. The first two are new files, the
// third replaces a real test's source.
const TEST_FIXTURES: ReadonlyArray<Fixture> = [
  {
    name: 'a write-side test that imports a read helper',
    path: WRITE_TEST_FIXTURE,
    source: `import { extractSharepointUrls } from '${specifierTo(WRITE_TEST_FIXTURE, LINK_EXTRACTOR)}';\nexport const f = extractSharepointUrls;\n`,
    reported: `${WRITE_TEST_FIXTURE} (write) imports ${LINK_EXTRACTOR} (read)`,
  },
  {
    name: 'a test that imports modules of two packages',
    path: MIXED_TEST_FIXTURE,
    source: `import type { AuthManager } from '${specifierTo(MIXED_TEST_FIXTURE, AUTH_LADDER)}';\nimport { extractSharepointUrls } from '${specifierTo(MIXED_TEST_FIXTURE, LINK_EXTRACTOR)}';\nexport type F = AuthManager;\nexport const f = extractSharepointUrls;\n`,
    reported: `${MIXED_TEST_FIXTURE} imports modules of auth and read`,
  },
  // The read registry is the single package's wiring until the move, then a read module: the
  // report names its owner after the path, so the fixtures match up to the path.
  {
    name: 'a test under core/ that imports the read registry',
    path: CORE_TEST_FIXTURE,
    source: `import { commands } from '${specifierTo(CORE_TEST_FIXTURE, REGISTRY)}';\nexport const f = commands;\n`,
    reported: `${CORE_TEST_FIXTURE} (core) imports ${REGISTRY} (`,
  },
  {
    name: 'a test under auth/ that imports the read registry',
    path: AUTH_TEST_FIXTURE,
    source: `import { commands } from '${specifierTo(AUTH_TEST_FIXTURE, REGISTRY)}';\nexport const f = commands;\n`,
    reported: `${AUTH_TEST_FIXTURE} (auth) imports ${REGISTRY} (`,
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
    reported: `${WRITE_FOLDER_TEST_FIXTURE} (write) imports ${REGISTRY} (`,
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
    name: 'a test outside write/ that imports write modules only',
    path: `${COMMANDS}/fixture-write-outside.test.ts`,
    source: "import { commands } from '@ask-marcel/office-write/index.ts';\nexport const f = commands;\n",
    reported: `${COMMANDS}/fixture-write-outside.test.ts (write) sits outside write/`,
  },
  {
    name: 'a test named after a write command that imports an auth module',
    path: MAIL_DRAFT_TEST,
    source: `import { parseTokenArgs } from '${specifierTo(MAIL_DRAFT_TEST, TOKEN_ARGS)}';\nexport const f = parseTokenArgs;\n`,
    reported: `${MAIL_DRAFT_TEST} (write) imports ${TOKEN_ARGS} (auth)`,
  },
];

// Must be reported by the rule that refuses a file in two ownership lists: a core file put in SINGLE_PACKAGE_FILES.
const OVERLAP_FIXTURE = { name: 'a file in two ownership lists', path: CORE_FIXTURE, reported: `${CORE_FIXTURE} (core, single)` } as const;

// Must be refused: the hoisted workspace would resolve it, and the check would not see its module.
const SPECIFIER_FIXTURE = {
  name: 'an import of a workspace package outside core, auth, read and write',
  path: LIST_DRIVES,
  specifier: '@ask-marcel/office-other/index.ts',
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
    if (test !== tests || (!test && /^((core|auth|read|write)\/)?src\/test-helpers\//.test(path))) continue;
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

// The package.json of a file's own package: its folder's, or the root's.
const packageJsonOf = (path: string): string => {
  const folder = FOLDERS.find((name) => path.startsWith(`${name}/`));
  return folder === undefined ? 'package.json' : `${folder}/package.json`;
};

const violations = (graph: Graph, owners: Owners): ReadonlyArray<string> => {
  const writeReach = [...owners.write].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || owners.write.has(path) || path === packageJsonOf(start)));
  const outsideWrite = [...graph.keys()].filter((start) => !owners.write.has(start));
  const writeReached = outsideWrite.flatMap((start) => crossing(graph, start, (path) => !owners.write.has(path)));
  const coreReach = [...owners.core].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path)));
  const outsideAuth = [...graph.keys()].filter((start) => !owners.auth.has(start));
  const authReached = outsideAuth.flatMap((start) => crossing(graph, start, (path) => !owners.auth.has(path)));
  const authReach = [...owners.auth].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || owners.auth.has(path) || path === packageJsonOf(start)));
  // A read module is every module that no other package owns (the root src/ and read/), never a package.json.
  const isRead = (path: string): boolean => !owners.core.has(path) && !owners.auth.has(path) && !owners.write.has(path) && !path.endsWith('package.json');
  const readReach = [...graph.keys()].filter(isRead).flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || isRead(path) || path === packageJsonOf(start)));
  return [...new Set([...writeReach, ...writeReached, ...coreReach, ...authReached, ...authReach, ...readReach])];
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
  if (path.startsWith('write/')) return 'write';
  if (path.startsWith('read/')) return 'read';
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
    if (own === 'write' && !path.startsWith('write/')) return [`${path} (write) sits outside write/: move it to write/src/`];
    // A test under core/ runs with core alone, and one under auth/ or write/ with
    // its package and core: the single wiring is not open to them.
    const closed = (target: string): boolean => PACKAGES.includes(owner(target)) || path.startsWith('core/') || (/^(auth|write)\//.test(path) && owner(target) !== 'core');
    const foreign = imports.filter((target) => closed(target) && owner(target) !== own);
    return [...new Set(foreign)].map((target) => `${path} (${own}) imports ${target} (${owner(target)})`);
  });

// The test graph keeps only the imports of modules under src/ (not package.json, not a package).
const buildTestGraph = (sources: Sources): Graph =>
  new Map([...buildGraph(sources)].map(([path, imports]) => [path, imports.filter((target) => /^((core|auth|read|write)\/)?src\//.test(target))]));

type CommandModule = { readonly meta?: { readonly effect?: unknown }; readonly execute?: unknown };

// The command folders: the root's until phase 6 moves the read commands to
// read's, and write's. A missing folder is left out (Bun.Glob.scanSync throws ENOENT on it).
const commandFolders = (): ReadonlyArray<string> => [COMMANDS, `read/${COMMANDS}`, `write/${COMMANDS}`].filter((folder) => existsSync(join(ROOT, folder)));

// The commands of the read registry and of the write package, split by effect.
const registryCommands = async (): Promise<{ readonly reads: ReadonlyArray<string>; readonly writes: ReadonlyArray<string> }> => {
  const reads: string[] = [];
  const writes: string[] = [];
  for (const folder of commandFolders()) {
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
  const core = glob([...CORE_FILES, ...CORE_READ_ONLY_FILES], 'core');
  const write = glob([], 'write');
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
    console.error(`check-package-imports: the ownership lists disagree with the registry or with each other. Write commands outside write/: ${unlisted.join(', ') || 'none'}. Files under write/ that are read commands or core: ${misplaced.join(', ') || 'none'}. Files in two lists: ${doubled.join(', ') || 'none'}. Read commands found: ${reads.length}.`);
    return 1;
  }
  const owners: Owners = { core, auth: lists.auth, write };
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
  console.log('check-package-imports: only write reaches write, and write reaches only core and write; core reaches only core; only auth reaches auth, and auth reaches only core and auth; read reaches only core and read; no test file imports a module of another package; every core test sits in core/, every auth test in auth/ and every write test in write/; every fixture is reported.');
  return 0;
};

process.exit(await main());
