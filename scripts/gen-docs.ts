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
 * each package ships the manifest of its own commands. The root `bun run
 * docs:gen` still joins the read and the write packages for the root docs
 * (see command-packages.ts):
 *
 *   bun run scripts/gen-docs.ts --package read/src/composition/single-package.ts \
 *     --package write/src/composition/write-package.ts \
 *     --manifest docs/commands.json --commands-doc docs/COMMANDS.md
 *
 * `--package` names a module that exports `commandPackage` (a CommandPackage),
 * or the export that `--export` names. It can come more than once: the
 * registries are joined, and the names come from the first package. Give
 * `--manifest`, `--commands-doc` or both. Paths are relative to the working
 * directory, and the manifest version is the one in its package.json.
 *
 * Markers: a commands doc that does not exist, a block with one marker only
 * (or with END before BEGIN), and a doc with no COMMANDS block pair when the
 * registry has commands each stop the run with exit 2. A doc with no LIFECYCLE
 * block pair is written: the root docs/COMMANDS.md has none.
 *
 * Pre-1.4.0 the auto-generated block lived inside README.md; the marketing
 * pass split docs out (README is the landing page, the deep tables moved
 * to docs/COMMANDS.md) so the README would stop ballooning past 500 lines.
 *
 */

import { resolve } from 'node:path';
import { parseArgs } from 'node:util';
import { fillNames } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import type { CommandManifest, CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';
import { renderLifecycleTable, renderReadmeTables } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';
import { buildManifest as buildHelpJsonManifest } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import { loadDocumentedPackage } from './command-packages.ts';

const USAGE =
  'usage: bun run ../scripts/gen-docs.ts --package <module exporting commandPackage> [--package <another>] [--export <name>] [--manifest <commands.json>] [--commands-doc <COMMANDS.md>]';

const { values: args } = parseArgs({
  options: { package: { type: 'string', multiple: true }, export: { type: 'string' }, manifest: { type: 'string' }, 'commands-doc': { type: 'string' } },
});
if (args.package === undefined || (args.manifest === undefined && args['commands-doc'] === undefined)) {
  process.stderr.write(`gen-docs: ${USAGE}\n`);
  process.exit(2);
}

const loaded = await loadDocumentedPackage(args.package, args.export);
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

const writeManifest = async (path: string): Promise<void> => {
  const manifest = buildManifest(await resolveGeneratedAt(path, buildManifest('')));
  await Bun.write(path, `${JSON.stringify(manifest, null, 2)}\n`);
  process.stderr.write(`gen-docs: wrote ${path} (${manifest.commands.length} commands)\n`);
};

const stop = (message: string): never => {
  process.stderr.write(`gen-docs: ${message}\n`);
  process.exit(2);
};

const beginMarker = (block: string): string => `<!-- AUTO-GENERATED-${block}:BEGIN -->`;
const endMarker = (block: string): string => `<!-- AUTO-GENERATED-${block}:END -->`;

type Span = { readonly begin: number; readonly end: number };

// Where the block pair is in the doc, or undefined when the doc has neither
// marker. One marker alone, or END before BEGIN, stops the run: a half pair
// would leave a block that no run updates.
const findBlock = (path: string, text: string, block: string): Span | undefined => {
  const begin = text.indexOf(beginMarker(block));
  const end = text.indexOf(endMarker(block));
  if (begin === -1 && end === -1) return undefined;
  if (begin === -1 || end === -1 || end < begin) return stop(`${path}: the ${block} block needs ${beginMarker(block)} and then ${endMarker(block)}.`);
  return { begin, end };
};

const replaceBlock = (path: string, text: string, block: string, body: string): string => {
  const span = findBlock(path, text, block);
  if (span === undefined) return text;
  return `${text.slice(0, span.begin + beginMarker(block).length)}\n\n${body}\n\n${text.slice(span.end)}`;
};

const rewriteCommandsDoc = async (path: string): Promise<void> => {
  const file = Bun.file(path);
  if (!(await file.exists())) stop(`${path} not found.`);
  const rawText = await file.text();
  const manifest = buildManifest('');
  // Keep the count sentences in sync with the registry so they can never drift
  // from the actual surface (they have before). N = registry command count; C =
  // distinct categories among them. The help-json total is derived from the real
  // manifest builder (registry + lifecycle entries), not hardcoded: it sat stale
  // at "179-vs-184" long after the registry had moved on. The first two forms
  // are those of the root docs/COMMANDS.md, the last two those of a package.
  const total = manifest.commands.length;
  const categoryCount = new Set(manifest.commands.map((c) => c.category)).size;
  const categories = `${categoryCount} ${categoryCount === 1 ? 'category' : 'categories'}`;
  const helpJsonTotal = buildHelpJsonManifest(commandPackage, version).commands.length;
  const text = rawText
    .replace(/All \d+ commands across \d+ categories/, `All ${total} commands across ${categoryCount} categories`)
    .replace(/manifest total \(\d+\), so a \d+-vs-\d+ gap/, `manifest total (${helpJsonTotal}), so a ${total}-vs-${helpJsonTotal} gap`)
    .replace(/The \d+ commands are in \d+ categor(?:y|ies)\./, `The ${total} commands are in ${categories}.`)
    .replace(
      /`help-json` lists \d+ entries: the \d+ commands and the \d+ lifecycle commands below\./,
      `\`help-json\` lists ${helpJsonTotal} entries: the ${total} commands and the ${commandPackage.lifecycle.length} lifecycle commands below.`
    );
  // A package with no command (auth) has no command tables.
  if (findBlock(path, text, 'COMMANDS') === undefined && total > 0) stop(`${path} has no ${beginMarker('COMMANDS')} block, and the registry has ${total} commands.`);
  const withLifecycle = replaceBlock(path, text, 'LIFECYCLE', renderLifecycleTable(commandPackage.lifecycle, fill));
  const next = total === 0 ? withLifecycle : replaceBlock(path, withLifecycle, 'COMMANDS', renderReadmeTables(manifest));
  if (next === rawText) {
    process.stderr.write(`gen-docs: ${path} already up-to-date\n`);
    return;
  }
  await Bun.write(path, next);
  process.stderr.write(`gen-docs: rewrote ${path}\n`);
};

if (args.manifest !== undefined) await writeManifest(args.manifest);
if (args['commands-doc'] !== undefined) await rewriteCommandsDoc(args['commands-doc']);
