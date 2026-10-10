#!/usr/bin/env bun
/*
 * The read guarantee on built code (package split, D9): a build of what read
 * will ship (the read graph and every command whose effect is `read`) holds no
 * write method and no upload session. While one package still holds both
 * graphs, its dist cannot answer that, so this check bundles the read side on
 * its own: an entry that imports `createReadGraph` and every read command
 * module, built the way `build:js` builds (every dependency external, so only
 * this repo's code is scanned).
 *
 * It looks for the markers only the write graph carries: a `"PUT"`, `"PATCH"`
 * or `"DELETE"` string, and `createUploadSession`. The ESLint rule in
 * eslint.config.js keeps read code from importing write-graph.ts; this proves
 * that nothing reaches it through another path either.
 *
 * Control: the same entry plus `createWriteGraph` must show all four markers,
 * so a check that can no longer see them fails instead of passing (rule 15.10:
 * a gate only ever seen green is a hypothesis).
 *
 * Exit codes:
 *   0  the read bundle has no marker, and the control shows every one
 *   1  the read bundle has a marker, the control misses one, or a build failed
 */

import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const ROOT = resolve(import.meta.dir, '..');
const COMMANDS = join(ROOT, 'src/use-cases/commands');
const READ_GRAPH = join(ROOT, 'src/infra/read-graph.ts');
const WRITE_GRAPH = join(ROOT, 'src/infra/write-graph.ts');

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

const dependencies = async (): Promise<ReadonlyArray<string>> => {
  const manifest = (await Bun.file(join(ROOT, 'package.json')).json()) as { readonly dependencies?: Record<string, string> };
  return Object.keys(manifest.dependencies ?? {});
};

const entrySource = (files: ReadonlyArray<string>, withWriteGraph: boolean): string =>
  [
    `export { createReadGraph } from ${JSON.stringify(READ_GRAPH)};`,
    ...files.map((file, index) => `export * as command${index} from ${JSON.stringify(file)};`),
    ...(withWriteGraph ? [`export { createWriteGraph } from ${JSON.stringify(WRITE_GRAPH)};`] : []),
  ].join('\n');

const bundle = async (dir: string, name: string, source: string, external: ReadonlyArray<string>): Promise<string | undefined> => {
  const entry = join(dir, `${name}.ts`);
  await Bun.write(entry, source);
  const built = await Bun.build({ entrypoints: [entry], target: 'node', format: 'esm', external: [...external] });
  if (!built.success) {
    console.error(`check-read-bundle: the ${name} build failed:\n${built.logs.map(String).join('\n')}`);
    return undefined;
  }
  const outputs = await Promise.all(built.outputs.map((output) => output.text()));
  return outputs.join('\n');
};

const found = (code: string): ReadonlyArray<string> => MARKERS.filter((marker) => marker.pattern.test(code)).map((marker) => marker.name);

const main = async (): Promise<number> => {
  const files = await readCommandFiles();
  const external = await dependencies();
  const dir = mkdtempSync(join(tmpdir(), 'read-bundle-'));
  try {
    const read = await bundle(dir, 'read', entrySource(files, false), external);
    const control = await bundle(dir, 'control', entrySource(files, true), external);
    if (read === undefined || control === undefined) return 1;
    const leaked = found(read);
    const seen = found(control);
    console.log(`check-read-bundle: the read graph and ${files.length} read commands, ${read.length} bytes built.`);
    if (leaked.length > 0) {
      console.error(`check-read-bundle: the read bundle holds write code: ${leaked.join(', ')}. A read module reaches the write graph; follow its imports.`);
      return 1;
    }
    if (seen.length !== MARKERS.length) {
      const missed = MARKERS.map((marker) => marker.name).filter((name) => !seen.includes(name));
      console.error(`check-read-bundle: the control bundle (with the write graph) misses ${missed.join(', ')}, so this check could no longer see write code.`);
      return 1;
    }
    console.log('check-read-bundle: no write method and no upload session in the read bundle; the control with the write graph shows all four.');
    return 0;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
};

process.exit(await main());
