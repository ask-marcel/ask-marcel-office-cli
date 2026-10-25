#!/usr/bin/env bun
/*
 * Documentation generator.
 *
 * Walks a package's registry, builds the JSON manifest, and rewrites the
 * generated blocks of the commands doc of the package:
 *
 *   AUTO-GENERATED-LIFECYCLE:BEGIN/END  the lifecycle commands of the package,
 *                                       each with its bin and the first
 *                                       sentence of its summary
 *   AUTO-GENERATED-COMMANDS:BEGIN/END   the per-category tables of the registry
 *
 * It also writes the counts of two sentences of the doc: "The N commands are
 * in C categories." and "`help-json` lists N entries: the M commands and the
 * L lifecycle commands below."
 *
 * The package and the paths are arguments, not imports, so one script serves
 * every package of the split (plan "Scripts ownership"). Each package folder
 * runs it from that folder:
 *
 *   read/:  bun run ../scripts/gen-docs.ts --package src/composition/single-package.ts --commands-doc COMMANDS.md
 *   write/: bun run ../scripts/gen-docs.ts --package src/composition/write-package.ts --commands-doc COMMANDS.md
 *   auth/:  bun run ../scripts/gen-docs.ts --package src/composition/auth-package.ts --export authPackage --commands-doc COMMANDS.md
 *
 * The builds of read/ and write/ run it with `--manifest dist/commands.json`, so
 * each package ships the manifest of its own commands.
 *
 * `--package` names a module that exports `commandPackage` (a CommandPackage),
 * or the export that `--export` names. A run takes one `--package` (a second
 * one is exit 2). Give `--manifest`, `--commands-doc` or both. Paths are
 * relative to the working directory, and the manifest version is the one in
 * its package.json.
 *
 * Markers: a commands doc that does not exist, a block with one marker only
 * (or with END before BEGIN), a doc with no LIFECYCLE block pair, and a doc
 * with no COMMANDS block pair when the registry has commands each stop the run
 * with exit 2. So no block can go dark.
 *
 *   --check     writes nothing. It renders the files in memory and compares
 *               them with the files on disk. The manifest is compared with
 *               its `generatedAt` masked, and a missing manifest is out of
 *               date. The `check:docs` script of each package folder runs it
 *               after the build.
 *   --selftest  renders a fresh doc of the package, then plants faults: a
 *               wrong count, a deleted table row, a changed lifecycle row and
 *               a changed manifest entry. Each fault must be out of date, the
 *               fresh doc and a manifest with another `generatedAt` must
 *               pass, and a missing manifest must be out of date. A doc with
 *               no LIFECYCLE pair (in the check and in the write mode), a
 *               lone BEGIN marker, an END marker before its BEGIN marker,
 *               and a doc with no COMMANDS pair must be refused (rule
 *               15.10). It needs `--package` only.
 *
 * Exit codes: 0 the files are up to date (or were written); 1 `--check`
 * found a file that is out of date, or the selftest failed; 2 a wrong
 * argument, a missing doc or a wrong marker.
 *
 * Pre-1.4.0 the auto-generated block lived inside README.md; the marketing
 * pass split docs out (README is the landing page, the deep tables moved
 * to docs/COMMANDS.md) so the README would stop ballooning past 500 lines.
 *
 */

import { resolve } from 'node:path';
import { parseArgs } from 'node:util';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { fillNames } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import type { CommandManifest, CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';
import { renderLifecycleTable, renderReadmeTables } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';
import { buildManifest as buildHelpJsonManifest } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import { loadPackage } from './command-packages.ts';

const USAGE =
  'usage: bun run ../scripts/gen-docs.ts --package <module exporting commandPackage> [--export <name>] [--manifest <commands.json>] [--commands-doc <COMMANDS.md>] [--check | --selftest]';

const { values: args } = parseArgs({
  options: {
    package: { type: 'string', multiple: true },
    export: { type: 'string' },
    manifest: { type: 'string' },
    'commands-doc': { type: 'string' },
    check: { type: 'boolean' },
    selftest: { type: 'boolean' },
  },
});
// One package for each run: parseArgs would keep only the last of two --package.
const [packagePath, ...extraPackages] = args.package ?? [];
if (packagePath === undefined || extraPackages.length > 0 || (args.selftest !== true && args.manifest === undefined && args['commands-doc'] === undefined)) {
  process.stderr.write(`gen-docs: ${USAGE}\n`);
  process.exit(2);
}

const loaded = await loadPackage(packagePath, args.export);
if (!loaded.ok) {
  process.stderr.write(`gen-docs: ${loaded.error} ${USAGE}\n`);
  process.exit(2);
}
const commandPackage = loaded.value;
const commands = commandPackage.registry;
const { version } = (await Bun.file(resolve('package.json')).json()) as { readonly version: string };

// The command references in meta carry the bin placeholders; the docs name the package's bins.
const fill = (text: string): string => fillNames(text, commandPackage);

const buildManifest = (generatedAt: string): CommandManifest => {
  const entries: CommandManifestEntry[] = [];
  for (const [name, cmd] of Object.entries(commands)) {
    const m = cmd.meta;
    entries.push({
      name,
      summary: fill(m.summary),
      category: m.category,
      graphMethod: m.graphMethod,
      graphPathTemplate: m.graphPathTemplate,
      graphDocsUrl: m.graphDocsUrl,
      options: m.options.map((option) => ({ ...option, description: fill(option.description) })),
      example: fill(m.example),
      ...(m.responseShape ? { responseShape: fill(m.responseShape) } : {}),
      ...(m.bodyTemplate ? { bodyTemplate: fill(m.bodyTemplate) } : {}),
      ...(m.pagination ? { pagination: m.pagination } : {}),
      ...(m.producesBytes ? { producesBytes: m.producesBytes } : {}),
      ...(m.producesMedia ? { producesMedia: m.producesMedia } : {}),
      effect: m.effect,
      ...(m.needsElevatedToken ? { needsElevatedToken: m.needsElevatedToken } : {}),
      ...(m.needsSubstrateToken ? { needsSubstrateToken: m.needsSubstrateToken } : {}),
      ...(m.stability ? { stability: m.stability } : {}),
    });
  }
  entries.sort((a, b) => a.name.localeCompare(b.name));
  return { package: commandPackage.packageName, version, generatedAt, commands: entries };
};

// Reuse the committed manifest's `generatedAt` when nothing but the timestamp
// would change — so a no-op run (e.g. the build prebuild) leaves the file byte-
// identical instead of dirtying git with a timestamp-only diff. A real content
// change (a new/edited command) stamps a fresh timestamp.
const resolveGeneratedAt = async (path: string, draft: CommandManifest): Promise<string> => {
  const file = Bun.file(path);
  if (!(await file.exists())) return new Date().toISOString();
  const prev = (await file.json()) as Partial<CommandManifest>;
  const sameContent = JSON.stringify({ ...draft, generatedAt: '' }) === JSON.stringify({ ...prev, generatedAt: '' });
  return sameContent && typeof prev.generatedAt === 'string' ? prev.generatedAt : new Date().toISOString();
};

const manifestText = (generatedAt: string): string => `${JSON.stringify(buildManifest(generatedAt), null, 2)}\n`;

const writeManifest = async (path: string): Promise<void> => {
  await Bun.write(path, manifestText(await resolveGeneratedAt(path, buildManifest(''))));
  process.stderr.write(`gen-docs: wrote ${path} (${Object.keys(commands).length} commands)\n`);
};

// A manifest is up to date when it equals a new one with its own `generatedAt`,
// so only the timestamp can differ.
const manifestIsFresh = (text: string): boolean => {
  const generatedAt = /^ {2}"generatedAt": "([^"]*)",$/m.exec(text)?.[1];
  return generatedAt !== undefined && text === manifestText(generatedAt);
};

// Why the manifest at `path` is out of date, or undefined when it is up to
// date. `text` is undefined when the file does not exist.
const manifestStale = (path: string, text: string | undefined): string | undefined => {
  if (text === undefined) return `${path} does not exist (run \`bun run build\`)`;
  return manifestIsFresh(text) ? undefined : `${path} (run \`bun run build\`)`;
};

const stop = (message: string): never => {
  process.stderr.write(`gen-docs: ${message}\n`);
  process.exit(2);
};

const beginMarker = (block: string): string => `<!-- AUTO-GENERATED-${block}:BEGIN -->`;
const endMarker = (block: string): string => `<!-- AUTO-GENERATED-${block}:END -->`;

// True when the doc has the block pair, false when it has neither marker. One
// marker alone, or END before BEGIN, is an error: a half pair would leave a
// block that no run updates.
const hasBlock = (path: string, text: string, block: string): Result<boolean, string> => {
  const begin = text.indexOf(beginMarker(block));
  const end = text.indexOf(endMarker(block));
  if (begin === -1 && end === -1) return ok(false);
  return begin === -1 || end < begin ? err(`${path}: the ${block} block needs ${beginMarker(block)} and then ${endMarker(block)}.`) : ok(true);
};

// Puts the body between the markers of a block pair that hasBlock found.
const fillBlock = (text: string, block: string, body: string): string =>
  `${text.slice(0, text.indexOf(beginMarker(block)) + beginMarker(block).length)}\n\n${body}\n\n${text.slice(text.indexOf(endMarker(block)))}`;

// The commands doc as a run writes it. A doc with no LIFECYCLE block pair is refused.
const renderCommandsDoc = (path: string, rawText: string): Result<string, string> => {
  const manifest = buildManifest('');
  // Keep the count sentences in sync with the registry so they can never drift
  // from the actual surface (they have before). N = registry command count; C =
  // distinct categories among them. The help-json total is derived from the real
  // manifest builder (registry + lifecycle entries), not hardcoded: it sat stale
  // at "179-vs-184" long after the registry had moved on.
  const total = manifest.commands.length;
  const categoryCount = new Set(manifest.commands.map((c) => c.category)).size;
  const categories = `${categoryCount} ${categoryCount === 1 ? 'category' : 'categories'}`;
  const helpJsonTotal = buildHelpJsonManifest(commandPackage, version).commands.length;
  const text = rawText
    .replace(/The \d+ commands are in \d+ categor(?:y|ies)\./, `The ${total} commands are in ${categories}.`)
    .replace(
      /`help-json` lists \d+ entries: the \d+ commands and the \d+ lifecycle commands below\./,
      `\`help-json\` lists ${helpJsonTotal} entries: the ${total} commands and the ${commandPackage.lifecycle.length} lifecycle commands below.`
    );
  const lifecycle = hasBlock(path, text, 'LIFECYCLE');
  if (!lifecycle.ok) return lifecycle;
  const tables = hasBlock(path, text, 'COMMANDS');
  if (!tables.ok) return tables;
  // A package with no command (auth) has no command tables.
  if (!tables.value && total > 0) return err(`${path} has no ${beginMarker('COMMANDS')} block, and the registry has ${total} commands.`);
  if (!lifecycle.value) return err(`${path} has no ${beginMarker('LIFECYCLE')} block.`);
  const withLifecycle = fillBlock(text, 'LIFECYCLE', renderLifecycleTable(commandPackage.lifecycle, fill));
  return ok(total === 0 ? withLifecycle : fillBlock(withLifecycle, 'COMMANDS', renderReadmeTables(manifest)));
};

const readCommandsDoc = async (path: string): Promise<string> => {
  const file = Bun.file(path);
  return (await file.exists()) ? file.text() : stop(`${path} not found.`);
};

const rewriteCommandsDoc = async (path: string): Promise<void> => {
  const rawText = await readCommandsDoc(path);
  const next = renderCommandsDoc(path, rawText);
  if (!next.ok) return stop(next.error);
  if (next.value === rawText) {
    process.stderr.write(`gen-docs: ${path} already up-to-date\n`);
    return;
  }
  await Bun.write(path, next.value);
  process.stderr.write(`gen-docs: rewrote ${path}\n`);
};

// Each file that is out of date, with the command that makes it again.
const staleFiles = async (manifestPath: string | undefined, docPath: string | undefined): Promise<ReadonlyArray<string>> => {
  const stale: string[] = [];
  if (manifestPath !== undefined) {
    const file = Bun.file(manifestPath);
    const problem = manifestStale(manifestPath, (await file.exists()) ? await file.text() : undefined);
    if (problem !== undefined) stale.push(problem);
  }
  if (docPath !== undefined) {
    const rawText = await readCommandsDoc(docPath);
    const next = renderCommandsDoc(docPath, rawText);
    if (!next.ok) return stop(next.error);
    if (next.value !== rawText) stale.push(`${docPath} (run \`bun run docs:gen\`)`);
  }
  return stale;
};

const check = async (): Promise<number> => {
  const stale = await staleFiles(args.manifest, args['commands-doc']);
  for (const file of stale) process.stderr.write(`gen-docs: out of date: ${file}\n`);
  if (stale.length === 0) process.stdout.write(`gen-docs: up to date: ${[args.manifest, args['commands-doc']].filter((p) => p !== undefined).join(', ')}\n`);
  return stale.length === 0 ? 0 : 1;
};

// The first table row after the BEGIN marker of a block.
const firstRow = (text: string, block: string): string =>
  text
    .slice(text.indexOf(beginMarker(block)))
    .split('\n')
    .find((line) => line.startsWith('| `')) ?? '';

const selftest = (): number => {
  const total = Object.keys(commands).length;
  if (total === 0) return stop('the selftest needs a package with commands.');
  const pair = (block: string): string => `${beginMarker(block)}\n${endMarker(block)}`;
  const head = 'The 0 commands are in 0 categories. `help-json` lists 0 entries: the 0 commands and the 0 lifecycle commands below.';
  const fresh = renderCommandsDoc('fresh', `${head}\n\n${pair('LIFECYCLE')}\n\n${pair('COMMANDS')}\n`);
  if (!fresh.ok) return stop(fresh.error);
  const isStale = (text: string): boolean => {
    const next = renderCommandsDoc('fault', text);
    return !next.ok || next.value !== text;
  };
  const doc = fresh.value;
  // A fault needs a row to change: with no row, its case fails.
  const tableRow = firstRow(doc, 'COMMANDS');
  const lifecycleRow = firstRow(doc, 'LIFECYCLE');
  const manifest = manifestText('2026-01-01T00:00:00.000Z');
  const refused = (text: string): boolean => !renderCommandsDoc('fault', text).ok;
  const withoutPair = (block: string): string => doc.replace(beginMarker(block), '').replace(endMarker(block), '');
  const swapped = doc.replace(beginMarker('COMMANDS'), '<swap>').replace(endMarker('COMMANDS'), beginMarker('COMMANDS')).replace('<swap>', endMarker('COMMANDS'));
  const cases = [
    { name: 'the fresh doc passes', passed: !isStale(doc) },
    { name: 'a wrong count is stale', passed: isStale(doc.replace(`The ${total} commands`, `The ${total + 1} commands`)) },
    { name: 'a deleted table row is stale', passed: tableRow !== '' && isStale(doc.replace(`${tableRow}\n`, '')) },
    { name: 'a changed lifecycle row is stale', passed: lifecycleRow !== '' && isStale(doc.replace(lifecycleRow, `${lifecycleRow} changed`)) },
    { name: 'a doc with no LIFECYCLE pair is refused (the check and the write mode share this render)', passed: refused(withoutPair('LIFECYCLE')) },
    { name: 'a lone BEGIN marker is refused', passed: refused(doc.replace(endMarker('LIFECYCLE'), '')) },
    { name: 'an END marker before its BEGIN marker is refused', passed: refused(swapped) },
    { name: 'a doc with no COMMANDS pair is refused when the registry has commands', passed: refused(withoutPair('COMMANDS')) },
    { name: 'a manifest with another generatedAt passes', passed: manifestIsFresh(manifest.replace('2026-01-01T00:00:00.000Z', '2026-02-02T00:00:00.000Z')) },
    { name: 'a changed manifest entry is stale', passed: !manifestIsFresh(manifest.replace('"summary": "', '"summary": "changed ')) },
    { name: 'a missing manifest is stale', passed: manifestStale('m.json', undefined) === 'm.json does not exist (run `bun run build`)' },
  ];
  const failed = cases.filter((c) => !c.passed);
  for (const c of failed) process.stderr.write(`gen-docs: SELFTEST FAILED: ${c.name}\n`);
  if (failed.length === 0) process.stdout.write(`gen-docs: selftest passed (${cases.map((c) => c.name).join('; ')}).\n`);
  return failed.length === 0 ? 0 : 1;
};

if (args.selftest === true) process.exit(selftest());
if (args.check === true) process.exit(await check());
if (args.manifest !== undefined) await writeManifest(args.manifest);
if (args['commands-doc'] !== undefined) await rewriteCommandsDoc(args['commands-doc']);
