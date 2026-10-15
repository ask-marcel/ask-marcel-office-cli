#!/usr/bin/env bun
/*
 * The package boundaries on the import graph (package split, step 15). Write
 * will bundle core and its own modules, never read's; read will bundle core and
 * its own modules, never write's. So, with type imports and dynamic imports
 * counted (a type still needs its module in the package):
 *
 *   - a write module (a write command or a write-only helper) reaches only core
 *     and write modules;
 *   - a read command and a core module never reach a write module.
 *
 * Ownership comes from eslint.config.js (WRITE_FILES, CORE_FILES,
 * CORE_READ_ONLY_FILES), the lists the core and read-only lint bans use, so a
 * file has one owner everywhere. The read commands are the registry's commands
 * whose effect is `read`; every other registry command must be a write module.
 *
 * Fixtures: each run also checks the tree with one file's source replaced by a
 * violating import, and each must be reported, so a check that can no longer
 * see an import fails instead of passing (rule 15.10: a gate only ever seen
 * green is a hypothesis).
 *
 * Exit codes:
 *   0  no boundary is crossed, and every fixture is reported
 *   1  a boundary is crossed, a fixture is missed, an import does not resolve,
 *      or the ownership lists disagree with the registry
 */

import { existsSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import ts from 'typescript';
import { CORE_FILES, CORE_READ_ONLY_FILES, WRITE_FILES } from '../eslint.config.js';

const ROOT = resolve(import.meta.dir, '..');
const COMMANDS = 'src/use-cases/commands';

type Sources = ReadonlyMap<string, string>;
type Graph = ReadonlyMap<string, ReadonlyArray<string>>;
type Owners = { readonly core: ReadonlySet<string>; readonly write: ReadonlySet<string>; readonly reads: ReadonlyArray<string> };
type Fixture = { readonly name: string; readonly path: string; readonly source: string; readonly reported: string };

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
    path: `${COMMANDS}/format-zod-error.ts`,
    source: "export const f = async (): Promise<unknown> => import('./create-reply-draft.ts');\n",
    reported: `${COMMANDS}/format-zod-error.ts reaches ${COMMANDS}/create-reply-draft.ts`,
  },
];

const glob = (patterns: ReadonlyArray<string>): ReadonlySet<string> =>
  new Set(patterns.flatMap((pattern) => [...new Bun.Glob(pattern).scanSync(ROOT)]).filter((path) => !path.endsWith('.test.ts')));

const readSources = async (): Promise<Sources> => {
  const sources = new Map<string, string>();
  for (const path of new Bun.Glob('src/**/*.ts').scanSync(ROOT)) {
    if (path.endsWith('.test.ts') || path.startsWith('src/test-helpers/')) continue;
    sources.set(path, await Bun.file(join(ROOT, path)).text());
  }
  return sources;
};

// The repo-relative file a relative specifier names (a module, or data such as
// package.json), or undefined for a package import. A relative specifier that
// names no file is an error.
const resolveSpecifier = (from: string, specifier: string): string | undefined => {
  if (!specifier.startsWith('.')) return undefined;
  const path = relative(ROOT, resolve(ROOT, dirname(from), specifier));
  const candidates = [path, `${path}.ts`, join(path, 'index.ts')];
  const found = candidates.find((candidate) => /\.[a-z]+$/.test(candidate) && existsSync(join(ROOT, candidate)));
  if (found === undefined) throw new Error(`${from}: the import '${specifier}' names no module`);
  return found;
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

const violations = (graph: Graph, owners: Owners): ReadonlyArray<string> => {
  const writeReach = [...owners.write].flatMap((start) => crossing(graph, start, (path) => owners.core.has(path) || owners.write.has(path)));
  const readReach = [...owners.reads, ...owners.core].flatMap((start) => crossing(graph, start, (path) => !owners.write.has(path)));
  return [...writeReach, ...readReach];
};

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
  const core = glob([...CORE_FILES, ...CORE_READ_ONLY_FILES]);
  const write = glob(WRITE_FILES);
  const { reads, writes } = await registryCommands();
  const unlisted = writes.filter((path) => !write.has(path));
  const misplaced = [...write].filter((path) => reads.includes(path) || core.has(path));
  if (unlisted.length > 0 || misplaced.length > 0 || reads.length === 0) {
    console.error(`check-package-imports: the ownership lists disagree with the registry. Write commands outside WRITE_FILES: ${unlisted.join(', ') || 'none'}. WRITE_FILES that are read commands or core: ${misplaced.join(', ') || 'none'}. Read commands found: ${reads.length}.`);
    return 1;
  }
  const owners: Owners = { core, write, reads };
  const sources = await readSources();
  const found = violations(buildGraph(sources), owners);
  const missed = FIXTURES.filter((fixture) => !violations(buildGraph(new Map([...sources, [fixture.path, fixture.source]])), owners).some((line) => line.startsWith(fixture.reported)));
  console.log(`check-package-imports: ${write.size} write modules, ${core.size} core modules, ${reads.length} read commands, ${sources.size} modules read.`);
  if (found.length > 0) console.error(`check-package-imports: a package boundary is crossed:\n  ${found.join('\n  ')}`);
  if (missed.length > 0) console.error(`check-package-imports: the check missed its fixture(s): ${missed.map((fixture) => fixture.name).join('; ')}.`);
  if (found.length > 0 || missed.length > 0) return 1;
  console.log('check-package-imports: write reaches only core and write; no read command and no core module reaches write; every fixture is reported.');
  return 0;
};

process.exit(await main());
