#!/usr/bin/env bun
/*
 * Documentation generator.
 *
 * Walks a package's registry, builds the JSON manifest, and rewrites the
 * per-category tables of the commands doc between the
 * AUTO-GENERATED-COMMANDS:BEGIN/END markers.
 *
 * The package and the two paths are arguments, not imports, so one script serves
 * every package of the split (plan "Scripts ownership"); `bun run docs:gen`
 * passes the single package's:
 *
 *   bun run scripts/gen-docs.ts --package src/composition/single-package.ts \
 *     --manifest docs/commands.json --commands-doc docs/COMMANDS.md
 *
 * `--package` names a module that exports `commandPackage` (a CommandPackage).
 * Paths are relative to the working directory, and the manifest version is the
 * one in its package.json.
 *
 * Pre-1.4.0 the auto-generated block lived inside README.md; the marketing
 * pass split docs out (README is the landing page, the deep tables moved
 * to docs/COMMANDS.md) so the README would stop ballooning past 500 lines.
 *
 * Run via `bun run docs:gen`. The build pipeline runs this BEFORE
 * `bun build` so `dist/commands.json` is always up-to-date.
 */

import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import type { CommandPackage } from '../src/use-cases/commands/command-package.ts';
import type { CommandManifest, CommandManifestEntry } from '../src/use-cases/commands/docs-render.ts';
import { renderReadmeTables } from '../src/use-cases/commands/docs-render.ts';
import { buildManifest as buildHelpJsonManifest } from '../src/use-cases/commands/docs-builders.ts';

const USAGE = 'usage: bun run scripts/gen-docs.ts --package <module exporting commandPackage> --manifest <commands.json> --commands-doc <COMMANDS.md>';

const { values: args } = parseArgs({ options: { package: { type: 'string' }, manifest: { type: 'string' }, 'commands-doc': { type: 'string' } } });
if (args.package === undefined || args.manifest === undefined || args['commands-doc'] === undefined) {
  process.stderr.write(`gen-docs: ${USAGE}\n`);
  process.exit(2);
}

const isCommandPackage = (value: unknown): value is CommandPackage =>
  typeof value === 'object' && value !== null && 'registry' in value && 'lifecycle' in value && 'packageName' in value;

const loaded: unknown = (await import(pathToFileURL(resolve(args.package)).href)).commandPackage;
if (!isCommandPackage(loaded)) {
  process.stderr.write(`gen-docs: ${args.package} exports no \`commandPackage\`. ${USAGE}\n`);
  process.exit(2);
}
const commandPackage = loaded;
const commands = commandPackage.registry;
const { version } = (await Bun.file(resolve('package.json')).json()) as { readonly version: string };

const COMMANDS_DOC_PATH = args['commands-doc'];
const MANIFEST_PATH = args.manifest;
const COMMANDS_DOC_BEGIN = '<!-- AUTO-GENERATED-COMMANDS:BEGIN -->';
const COMMANDS_DOC_END = '<!-- AUTO-GENERATED-COMMANDS:END -->';

const buildManifest = (generatedAt: string): CommandManifest => {
  const entries: CommandManifestEntry[] = [];
  for (const [name, cmd] of Object.entries(commands)) {
    const m = cmd.meta;
    entries.push({
      name,
      summary: m.summary,
      category: m.category,
      graphMethod: m.graphMethod,
      graphPathTemplate: m.graphPathTemplate,
      graphDocsUrl: m.graphDocsUrl,
      options: m.options,
      example: m.example,
      ...(m.responseShape ? { responseShape: m.responseShape } : {}),
      ...(m.bodyTemplate ? { bodyTemplate: m.bodyTemplate } : {}),
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
const resolveGeneratedAt = async (draft: CommandManifest): Promise<string> => {
  const file = Bun.file(MANIFEST_PATH);
  if (!(await file.exists())) return new Date().toISOString();
  const prev = (await file.json()) as Partial<CommandManifest>;
  const sameContent = JSON.stringify({ ...draft, generatedAt: '' }) === JSON.stringify({ ...prev, generatedAt: '' });
  return sameContent && typeof prev.generatedAt === 'string' ? prev.generatedAt : new Date().toISOString();
};

const writeManifest = async (manifest: CommandManifest): Promise<void> => {
  await Bun.write(MANIFEST_PATH, `${JSON.stringify(manifest, null, 2)}\n`);
  process.stderr.write(`gen-docs: wrote ${MANIFEST_PATH} (${manifest.commands.length} commands)\n`);
};

const rewriteCommandsDoc = async (manifest: CommandManifest): Promise<void> => {
  const file = Bun.file(COMMANDS_DOC_PATH);
  if (!(await file.exists())) {
    process.stderr.write(`gen-docs: ${COMMANDS_DOC_PATH} not found — skipping rewrite\n`);
    return;
  }
  const rawText = await file.text();
  // Keep the headline "All N commands across M categories" in sync with the
  // registry so it can never drift from the actual surface (it has before).
  // N = registry command count; M = distinct categories among them (the
  // lifecycle commands live in their own hand-written table, not the registry).
  // The help-json total in the same sentence is derived from the real manifest
  // builder (registry + lifecycle entries), not hardcoded: it sat stale at
  // "179-vs-184" long after the registry had moved on.
  const categoryCount = new Set(manifest.commands.map((c) => c.category)).size;
  const helpJsonTotal = buildHelpJsonManifest(commandPackage, version).commands.length;
  const text = rawText
    .replace(/All \d+ commands across \d+ categories/, `All ${manifest.commands.length} commands across ${categoryCount} categories`)
    .replace(/manifest total \(\d+\), so a \d+-vs-\d+ gap/, `manifest total (${helpJsonTotal}), so a ${manifest.commands.length}-vs-${helpJsonTotal} gap`);
  const begin = text.indexOf(COMMANDS_DOC_BEGIN);
  const end = text.indexOf(COMMANDS_DOC_END);
  if (begin === -1 || end === -1 || end < begin) {
    process.stderr.write(`gen-docs: ${COMMANDS_DOC_PATH} markers not found (${COMMANDS_DOC_BEGIN} / ${COMMANDS_DOC_END}) — skipping rewrite\n`);
    return;
  }
  const before = text.slice(0, begin + COMMANDS_DOC_BEGIN.length);
  const after = text.slice(end);
  const generated = `\n\n${renderReadmeTables(manifest)}\n\n`;
  const next = `${before}${generated}${after}`;
  if (next === rawText) {
    process.stderr.write(`gen-docs: ${COMMANDS_DOC_PATH} already up-to-date\n`);
    return;
  }
  await Bun.write(COMMANDS_DOC_PATH, next);
  process.stderr.write(`gen-docs: rewrote ${COMMANDS_DOC_PATH} command tables\n`);
};

const draft = buildManifest('');
const manifest = buildManifest(await resolveGeneratedAt(draft));
await writeManifest(manifest);
await rewriteCommandsDoc(manifest);
