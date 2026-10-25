// The auth cases of the docs-builder suite: the pins on the lifecycle entries of `login`,
// `logout` and `status`, which only the auth package will list. Moved out of docs.test.ts
// with their text unchanged (package split, step 16), so the auth package takes this file
// whole; the helpers are copies of that file's.
import { describe, expect, it } from 'bun:test';
import { commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import type { CommandPackage } from '@ask-marcel/office-core/use-cases/commands/command-package.ts';
import type { Command } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { buildManifest, buildTerseManifest, renderSingleCommand } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import type { CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';
import { AUTH_LIFECYCLE_ENTRIES as LIFECYCLE_ENTRIES } from './auth-lifecycle-entries.ts';

const fakeCmd = (overrides: Partial<Command['meta']> = {}): Command => ({
  schema: { _: 'fake' } as never,
  execute: async () => ({ ok: true, value: undefined }),
  meta: {
    summary: 'fake summary',
    category: 'drive',
    graphMethod: 'GET',
    graphPathTemplate: '/fake',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/fake',
    options: [],
    example: 'ask-marcel-office fake',
    effect: 'read',
    ...overrides,
  },
});

const LIFECYCLE_NAMES = ['docs', 'help-json', 'login', 'logout', 'mcp', 'status'] as const;

// A test registry with the auth package's lifecycle entries, as the docs builders take it.
const withLifecycle = (registry: Readonly<Record<string, Command>>, packageName = 'fake-pkg'): Pick<CommandPackage, 'registry' | 'lifecycle' | 'packageName'> => ({
  registry,
  lifecycle: LIFECYCLE_ENTRIES,
  packageName,
});

describe('buildManifest', () => {
  it('builds a manifest with package name, version, generatedAt, and registry+lifecycle commands sorted alphabetically', () => {
    const registry: Readonly<Record<string, Command>> = { 'list-zebra': fakeCmd(), 'list-apple': fakeCmd() };
    const manifest = buildManifest(withLifecycle(registry), '0.0.1', () => new Date('2026-04-30T12:00:00Z'));
    expect(manifest.package).toBe('fake-pkg');
    expect(manifest.version).toBe('0.0.1');
    expect(manifest.generatedAt).toBe('2026-04-30T12:00:00.000Z');
    expect(manifest.commands.map((c) => c.name)).toEqual(['docs', 'help-json', 'list-apple', 'list-zebra', 'login', 'logout', 'mcp', 'status']);
  });

  it('marks every lifecycle entry with category `lifecycle` so consumers can filter them', () => {
    const manifest = buildManifest(withLifecycle({}), '0.0.1');
    const lifecycle = manifest.commands.filter((c) => LIFECYCLE_NAMES.includes(c.name as (typeof LIFECYCLE_NAMES)[number]));
    expect(lifecycle).toHaveLength(LIFECYCLE_NAMES.length);
    for (const entry of lifecycle) expect(entry.category).toBe('lifecycle');
  });
});

describe('buildTerseManifest — discovery view', () => {
  it('still includes lifecycle entries with their canonical summaries so a discovery-mode consumer sees login/logout/docs/help-json', () => {
    const manifest = buildTerseManifest(withLifecycle({}), '0.0.1');
    const names = manifest.commands.map((c) => c.name);
    for (const lifecycle of LIFECYCLE_NAMES) expect(names).toContain(lifecycle);
  });
});

describe('renderSingleCommand', () => {
  it('returns Markdown for a lifecycle command (login/logout/docs/help-json) even when the registry is empty', () => {
    const result = renderSingleCommand(withLifecycle({}), 'login');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('# `login`');
      expect(result.value).toContain('Authenticate against Microsoft Graph');
    }
  });

  it('returns unknown_command with the alphabetically merged registry+lifecycle list when the command is missing', () => {
    const registry: Readonly<Record<string, Command>> = { 'list-zebra': fakeCmd(), 'list-apple': fakeCmd() };
    const result = renderSingleCommand(withLifecycle(registry), 'list-banana');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'unknown_command') {
      expect(result.error.name).toBe('list-banana');
      expect(result.error.available).toEqual(['docs', 'help-json', 'list-apple', 'list-zebra', 'login', 'logout', 'mcp', 'status']);
    }
  });
});

describe('lifecycle entries — the invariants the registry sweep never reaches', () => {
  // login/status/logout/docs/help-json/mcp have no command file, so they are absent
  // from the `commands` registry and skip every invariant meta.test.ts applies to
  // the other 186. buildManifest merging them in is the only public route to them.
  const lifecycle = buildManifest(withLifecycle({}), '0.0.1').commands;
  const byName = (name: string): CommandManifestEntry | undefined => lifecycle.find((c) => c.name === name);

  it('exposes exactly the six lifecycle commands', () => {
    expect(lifecycle.map((c) => c.name)).toEqual([...LIFECYCLE_NAMES]);
  });

  for (const name of LIFECYCLE_NAMES) {
    it(`\`${name}\` carries the same populated meta a registered command must`, () => {
      const entry = byName(name);
      expect(entry).toBeDefined();
      if (!entry) return;
      expect(entry.summary.trim().length).toBeGreaterThan(40);
      expect(entry.graphPathTemplate.trim().length).toBeGreaterThan(0);
      expect(['GET', 'POST', 'PATCH', 'PUT', 'DELETE']).toContain(entry.graphMethod);
      expect(entry.graphDocsUrl.startsWith('https://')).toBe(true);
      // The example is what an LLM copies verbatim, so it has to be a real invocation:
      // the bin of the package that holds the command (command-directory.ts), or
      // `{bin}` for one every package has.
      const role = commandRole(name);
      const bin = role === undefined ? '{bin}' : `{${role}-bin}`;
      expect(entry.example.startsWith(`${bin} ${name}`)).toBe(true);
      expect(entry.responseShape?.trim().length ?? 0).toBeGreaterThan(0);
      for (const opt of entry.options) {
        expect(opt.name.trim().length).toBeGreaterThan(0);
        expect(opt.key.trim().length).toBeGreaterThan(0);
        expect(opt.description.trim().length).toBeGreaterThan(0);
      }
      for (const arg of entry.positionalArguments ?? []) {
        expect(arg.name.trim().length).toBeGreaterThan(0);
        expect(arg.description.trim().length).toBeGreaterThan(0);
      }
    });
  }

  it('keeps `login --force` declared, and optional — a required flag would break bare `login`', () => {
    const force = byName('login')?.options.find((o) => o.name === 'force');
    expect(force?.key).toBe('force');
    expect(force?.required).toBe(false);
  });
});
