#!/usr/bin/env bun
/*
 * The read guarantee on built code (package split, D9): a build of what read
 * will ship (the read graph and every command whose effect is `read`) holds no
 * write method and no upload session. While one package still holds both
 * graphs, its dist cannot answer that, so this check bundles the read side on
 * its own: an entry that imports `createReadGraph` and every read command
 * module, built the way read's build does (the dependencies of
 * read/package.json external, every other module inlined).
 *
 * It looks for the markers only the write graph carries: a `"PUT"`, `"PATCH"`
 * or `"DELETE"` string, and `createUploadSession`. The ESLint rule in
 * read/eslint.config.js keeps read code from importing write-graph.ts; this
 * proves that nothing reaches it through another path either.
 *
 * The real read entries are read bundles too since the root cut (package
 * split, phase 5): read/src/main.ts (the bin) and read/src/index.ts (the
 * library), built the same way, must hold no marker either.
 *
 * Controls: the synthetic entry plus `createWriteGraph`, and the real entry of
 * the write bin (write/src/main.ts, with the dependencies of
 * write/package.json external), must each show all four markers, so a
 * check that can no longer see them fails instead of passing (rule 15.10: a
 * gate only ever seen green is a hypothesis).
 *
 * Exit codes:
 *   0  no read bundle has a marker, and each control shows every one
 *   1  a read bundle has a marker, a control misses one, or a build failed
 */

import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const ROOT = resolve(import.meta.dir, '..');
// Read's source folder (package split, phase 6).
const READ_SRC = 'read/src';
const COMMANDS = join(ROOT, READ_SRC, 'use-cases/commands');
// The two graphs are core's (package split, phase 3).
const READ_GRAPH = join(ROOT, 'core/src/infra/read-graph.ts');
const WRITE_GRAPH = join(ROOT, 'core/src/infra/write-graph.ts');
// The real entries: read's and the write bin (a control).
const READ_ENTRIES: ReadonlyArray<string> = [`${READ_SRC}/main.ts`, `${READ_SRC}/index.ts`];
const WRITE_ENTRY = 'write/src/main.ts';

type Marker = { readonly name: string; readonly pattern: RegExp };

const MARKERS: ReadonlyArray<Marker> = [
  { name: "'PUT'", pattern: /(["'`])PUT\1/ },
  { name: "'PATCH'", pattern: /(["'`])PATCH\1/ },
  { name: "'DELETE'", pattern: /(["'`])DELETE\1/ },
  { name: 'createUploadSession', pattern: /createUploadSession/ },
];

type CommandModule = { readonly meta?: { readonly effect?: unknown }; readonly execute?: unknown };

// Every module in the commands folder that is a command whose effect is read.
const readCommandFiles = async (): Promise<ReadonlyArray<string>> => {
  const files = [...new Bun.Glob('*.ts').scanSync(COMMANDS)].filter((file) => !file.endsWith('.test.ts')).toSorted((a, b) => a.localeCompare(b));
  const reads: string[] = [];
  for (const file of files) {
    const module = (await import(join(COMMANDS, file))) as CommandModule;
    if (typeof module.execute === 'function' && module.meta?.effect === 'read') reads.push(join(COMMANDS, file));
  }
  return reads;
};

// The externals of a package's bundles: the keys of its dependencies (scripts/build-package.ts).
const dependencies = async (folder: string): Promise<ReadonlyArray<string>> => {
  const manifest = (await Bun.file(join(ROOT, folder, 'package.json')).json()) as { readonly dependencies?: Record<string, string> };
  return Object.keys(manifest.dependencies ?? {});
};

const entrySource = (files: ReadonlyArray<string>, withWriteGraph: boolean): string =>
  [
    `export { createReadGraph } from ${JSON.stringify(READ_GRAPH)};`,
    ...files.map((file, index) => `export * as command${index} from ${JSON.stringify(file)};`),
    ...(withWriteGraph ? [`export { createWriteGraph } from ${JSON.stringify(WRITE_GRAPH)};`] : []),
  ].join('\n');

const build = async (name: string, entry: string, external: ReadonlyArray<string>): Promise<string | undefined> => {
  const built = await Bun.build({ entrypoints: [entry], target: 'node', format: 'esm', external: [...external] });
  if (!built.success) {
    console.error(`check-read-bundle: the ${name} build failed:\n${built.logs.map(String).join('\n')}`);
    return undefined;
  }
  const outputs = await Promise.all(built.outputs.map((output) => output.text()));
  return outputs.join('\n');
};

const bundle = async (dir: string, name: string, source: string, external: ReadonlyArray<string>): Promise<string | undefined> => {
  const entry = join(dir, `${name}.ts`);
  await Bun.write(entry, source);
  return build(name, entry, external);
};

const found = (code: string): ReadonlyArray<string> => MARKERS.filter((marker) => marker.pattern.test(code)).map((marker) => marker.name);

const main = async (): Promise<number> => {
  const files = await readCommandFiles();
  const external = await dependencies('read');
  const dir = mkdtempSync(join(tmpdir(), 'read-bundle-'));
  try {
    const read = await bundle(dir, 'read', entrySource(files, false), external);
    const control = await bundle(dir, 'control', entrySource(files, true), external);
    const entries = await Promise.all(READ_ENTRIES.map(async (entry) => [entry, await build(entry, join(ROOT, entry), external)] as const));
    const writeControl = await build(WRITE_ENTRY, join(ROOT, WRITE_ENTRY), await dependencies('write'));
    if (read === undefined || control === undefined || writeControl === undefined || entries.some(([, code]) => code === undefined)) return 1;
    console.log(`check-read-bundle: the read graph and ${files.length} read commands, ${read.length} bytes built; ${READ_ENTRIES.join(' and ')} built as read bundles.`);
    const leaks = [['the read bundle', read] as const, ...entries.map(([entry, code]) => [`the bundle of ${entry}`, code ?? ''] as const)].flatMap(([name, code]) => {
      const leaked = found(code);
      return leaked.length === 0 ? [] : [`check-read-bundle: ${name} holds write code: ${leaked.join(', ')}. A read module reaches the write graph; follow its imports.`];
    });
    const blind = [['the control bundle (with the write graph)', control] as const, [`the control bundle of ${WRITE_ENTRY} (the write bin)`, writeControl] as const].flatMap(
      ([name, code]) => {
        const seen = found(code);
        const missed = MARKERS.map((marker) => marker.name).filter((marker) => !seen.includes(marker));
        return missed.length === 0 ? [] : [`check-read-bundle: ${name} misses ${missed.join(', ')}, so this check could no longer see write code.`];
      }
    );
    for (const problem of [...leaks, ...blind]) console.error(problem);
    if (leaks.length + blind.length > 0) return 1;
    console.log(
      `check-read-bundle: no write method and no upload session in the read bundle, ${READ_ENTRIES.join(' or ')}; the control with the write graph and ${WRITE_ENTRY} show all four.`
    );
    return 0;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
};

process.exit(await main());
