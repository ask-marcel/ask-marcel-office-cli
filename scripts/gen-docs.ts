#!/usr/bin/env bun
/*
 * Documentation generator.
 *
 * Walks the `commands` registry, builds the JSON manifest at
 * `docs/commands.json`, and rewrites the per-category tables inside
 * `docs/COMMANDS.md` between the AUTO-GENERATED-COMMANDS:BEGIN/END markers.
 *
 * Pre-1.4.0 the auto-generated block lived inside README.md; the marketing
 * pass split docs out (README is the landing page, the deep tables moved
 * to docs/COMMANDS.md) so the README would stop ballooning past 500 lines.
 *
 * Run via `bun run docs:gen`. The build pipeline runs this BEFORE
 * `bun build` so `dist/commands.json` is always up-to-date.
 */

import pkg from '../package.json' with { type: 'json' };
import { commands } from '../src/use-cases/commands/index.ts';
import type { CommandManifest, CommandManifestEntry } from '../src/use-cases/commands/docs-render.ts';
import { renderReadmeTables } from '../src/use-cases/commands/docs-render.ts';

const COMMANDS_DOC_PATH = 'docs/COMMANDS.md';
const MANIFEST_PATH = 'docs/commands.json';
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
      ...(m.stability ? { stability: m.stability } : {}),
    });
  }
  entries.sort((a, b) => a.name.localeCompare(b.name));
  return { package: pkg.name, version: pkg.version, generatedAt, commands: entries };
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
  const text = await file.text();
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
  if (next === text) {
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
