#!/usr/bin/env bun
/*
 * Documentation generator.
 *
 * Walks the `commands` registry, builds the JSON manifest at
 * `docs/commands.json`, and rewrites the per-category tables inside the
 * README between the AUTO-GENERATED-COMMANDS:BEGIN/END markers.
 *
 * Run via `bun run docs:gen`. The build pipeline runs this BEFORE
 * `bun build` so `dist/commands.json` is always up-to-date.
 */

import pkg from '../package.json' with { type: 'json' };
import { commands } from '../src/use-cases/commands/index.ts';
import type { CommandManifest, CommandManifestEntry } from '../src/use-cases/commands/docs-render.ts';
import { renderReadmeTables } from '../src/use-cases/commands/docs-render.ts';

const README_PATH = 'README.md';
const MANIFEST_PATH = 'docs/commands.json';
const README_BEGIN = '<!-- AUTO-GENERATED-COMMANDS:BEGIN -->';
const README_END = '<!-- AUTO-GENERATED-COMMANDS:END -->';

const buildManifest = (): CommandManifest => {
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
  return {
    package: pkg.name,
    version: pkg.version,
    generatedAt: new Date().toISOString(),
    commands: entries,
  };
};

const writeManifest = async (manifest: CommandManifest): Promise<void> => {
  await Bun.write(MANIFEST_PATH, `${JSON.stringify(manifest, null, 2)}\n`);
  process.stderr.write(`gen-docs: wrote ${MANIFEST_PATH} (${manifest.commands.length} commands)\n`);
};

const rewriteReadme = async (manifest: CommandManifest): Promise<void> => {
  const file = Bun.file(README_PATH);
  if (!(await file.exists())) {
    process.stderr.write(`gen-docs: ${README_PATH} not found — skipping README rewrite\n`);
    return;
  }
  const text = await file.text();
  const begin = text.indexOf(README_BEGIN);
  const end = text.indexOf(README_END);
  if (begin === -1 || end === -1 || end < begin) {
    process.stderr.write(`gen-docs: README markers not found (${README_BEGIN} / ${README_END}) — skipping README rewrite\n`);
    return;
  }
  const before = text.slice(0, begin + README_BEGIN.length);
  const after = text.slice(end);
  const generated = `\n\n${renderReadmeTables(manifest)}\n\n`;
  const next = `${before}${generated}${after}`;
  if (next === text) {
    process.stderr.write(`gen-docs: README already up-to-date\n`);
    return;
  }
  await Bun.write(README_PATH, next);
  process.stderr.write(`gen-docs: rewrote ${README_PATH} command tables\n`);
};

const manifest = buildManifest();
await writeManifest(manifest);
await rewriteReadme(manifest);
