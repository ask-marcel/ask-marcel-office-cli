import { describe, expect, it } from 'bun:test';
import type { Command } from './command-types.ts';
import { buildManifest, buildTerseManifest, filterManifestByCategory, renderSingleCommand } from './docs.ts';
import { commands } from './index.ts';

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
    ...overrides,
  },
});

const LIFECYCLE_NAMES = ['docs', 'help-json', 'login', 'logout', 'update'] as const;

describe('buildManifest', () => {
  it('builds a manifest with package name, version, generatedAt, and registry+lifecycle commands sorted alphabetically', () => {
    const registry: Readonly<Record<string, Command>> = { 'list-zebra': fakeCmd(), 'list-apple': fakeCmd() };
    const manifest = buildManifest(registry, 'fake-pkg', '0.0.1', () => new Date('2026-04-30T12:00:00Z'));
    expect(manifest.package).toBe('fake-pkg');
    expect(manifest.version).toBe('0.0.1');
    expect(manifest.generatedAt).toBe('2026-04-30T12:00:00.000Z');
    expect(manifest.commands.map((c) => c.name)).toEqual(['docs', 'help-json', 'list-apple', 'list-zebra', 'login', 'logout', 'update']);
  });

  it('marks every lifecycle entry with category `lifecycle` so consumers can filter them', () => {
    const manifest = buildManifest({}, 'fake-pkg', '0.0.1');
    const lifecycle = manifest.commands.filter((c) => LIFECYCLE_NAMES.includes(c.name as (typeof LIFECYCLE_NAMES)[number]));
    expect(lifecycle).toHaveLength(LIFECYCLE_NAMES.length);
    for (const entry of lifecycle) expect(entry.category).toBe('lifecycle');
  });

  it('omits responseShape when the source registry meta does not provide one', () => {
    const registry: Readonly<Record<string, Command>> = { 'aaa-foo': fakeCmd() };
    const manifest = buildManifest(registry, 'fake-pkg', '0.0.1');
    const fooEntry = manifest.commands.find((c) => c.name === 'aaa-foo');
    expect(fooEntry).not.toHaveProperty('responseShape');
  });

  it('keeps responseShape when the source registry meta provides one', () => {
    const registry: Readonly<Record<string, Command>> = { 'aaa-foo': fakeCmd({ responseShape: 'single thing' }) };
    const manifest = buildManifest(registry, 'fake-pkg', '0.0.1');
    const fooEntry = manifest.commands.find((c) => c.name === 'aaa-foo');
    expect(fooEntry?.responseShape).toBe('single thing');
  });

  it('serializes producesBytes / producesMedia / mutates when set, and omits them otherwise (F-01 — help-json must mirror the meta, not silently drop byte/media/write flags)', () => {
    const registry: Readonly<Record<string, Command>> = {
      'aaa-plain': fakeCmd(),
      'aaa-bytes': fakeCmd({ producesBytes: true }),
      'aaa-media': fakeCmd({ producesMedia: true }),
      'aaa-write': fakeCmd({ graphMethod: 'PATCH', mutates: true }),
    };
    const manifest = buildManifest(registry, 'fake-pkg', '0.0.1');
    const byName = (n: string): (typeof manifest.commands)[number] | undefined => manifest.commands.find((c) => c.name === n);
    expect(byName('aaa-plain')).not.toHaveProperty('producesBytes');
    expect(byName('aaa-plain')).not.toHaveProperty('producesMedia');
    expect(byName('aaa-plain')).not.toHaveProperty('mutates');
    expect(byName('aaa-bytes')?.producesBytes).toBe(true);
    expect(byName('aaa-media')?.producesMedia).toBe(true);
    expect(byName('aaa-write')?.mutates).toBe(true);
  });

  it('serializes commandAliases (back-compat names) when set, and omits the key otherwise', () => {
    const registry: Readonly<Record<string, Command>> = {
      'aaa-renamed': fakeCmd({ commandAliases: ['aaa-old-name'] }),
      'aaa-plain': fakeCmd(),
    };
    const manifest = buildManifest(registry, 'fake-pkg', '0.0.1');
    expect(manifest.commands.find((c) => c.name === 'aaa-renamed')?.commandAliases).toEqual(['aaa-old-name']);
    expect(manifest.commands.find((c) => c.name === 'aaa-plain')).not.toHaveProperty('commandAliases');
  });

  it('uses the real `new Date()` when no clock injector is given', () => {
    const before = Date.now();
    const manifest = buildManifest({ foo: fakeCmd() }, 'fake-pkg', '0.0.1');
    const after = Date.now();
    const generatedAt = new Date(manifest.generatedAt).getTime();
    expect(generatedAt).toBeGreaterThanOrEqual(before);
    expect(generatedAt).toBeLessThanOrEqual(after);
  });
});

describe('buildTerseManifest — discovery view', () => {
  it('strips every per-command field other than name/summary/category from each entry', () => {
    const registry: Readonly<Record<string, Command>> = {
      'list-foo': fakeCmd({ summary: 'lists foos', responseShape: 'array of foos', bodyTemplate: '{ "x": 1 }', pagination: true }),
    };
    const manifest = buildTerseManifest(registry, 'fake-pkg', '0.0.1', () => new Date('2026-04-30T12:00:00Z'));
    const foo = manifest.commands.find((c) => c.name === 'list-foo');
    expect(foo).toEqual({ name: 'list-foo', summary: 'lists foos', category: 'drive' });
  });

  it('compacts each terse summary to its first sentence (keeps the per-category discovery view token-cheap)', () => {
    const registry: Readonly<Record<string, Command>> = {
      'list-foo': fakeCmd({ summary: 'Lists the foos. A second sentence with extra detail that terse should drop.' }),
    };
    const manifest = buildTerseManifest(registry, 'fake-pkg', '0.0.1');
    const foo = manifest.commands.find((c) => c.name === 'list-foo');
    expect(foo?.summary).toBe('Lists the foos.');
  });

  it('still includes lifecycle entries with their canonical summaries so a discovery-mode consumer sees login/logout/update/docs/help-json', () => {
    const manifest = buildTerseManifest({}, 'fake-pkg', '0.0.1');
    const names = manifest.commands.map((c) => c.name);
    for (const lifecycle of LIFECYCLE_NAMES) expect(names).toContain(lifecycle);
  });

  it('keeps the `stability` tag on terse entries so an LLM sees the experimental marker at discovery time (— no second full-manifest fetch needed)', () => {
    const registry: Readonly<Record<string, Command>> = {
      'list-stable-thing': fakeCmd(),
      'list-experimental-thing': fakeCmd({ stability: 'experimental' }),
    };
    const manifest = buildTerseManifest(registry, 'fake-pkg', '0.0.1');
    const stable = manifest.commands.find((c) => c.name === 'list-stable-thing');
    const experimental = manifest.commands.find((c) => c.name === 'list-experimental-thing');
    expect(stable?.stability).toBeUndefined();
    expect(experimental?.stability).toBe('experimental');
  });

  it('shrinks the wire payload substantially versus the full manifest (regression guard on the discovery-view contract)', () => {
    const registry: Readonly<Record<string, Command>> = Object.fromEntries(
      Array.from({ length: 50 }, (_, i) => [`list-thing-${i}`, fakeCmd({ summary: 'x'.repeat(400), responseShape: 'y'.repeat(400), bodyTemplate: 'z'.repeat(400) })])
    );
    const full = JSON.stringify(buildManifest(registry, 'fake-pkg', '0.0.1'));
    const terse = JSON.stringify(buildTerseManifest(registry, 'fake-pkg', '0.0.1'));
    // Terse should be at least 50% smaller — the heavy responseShape /
    // bodyTemplate fields drop out entirely on every entry.
    expect(terse.length).toBeLessThan(full.length / 2);
  });

  it('keeps the real help-json sizes within their documented budget — a tripwire so the "~440 KB full / ~31 KB terse / ~6 KB per category" hints (README/USAGE/cli/error-hints) cannot silently go stale as commands are added', () => {
    const stamp = (): Date => new Date(0);
    const full = JSON.stringify(buildManifest(commands, 'ask-marcel-office-cli', '0.0.0', stamp)).length;
    const terseAll = JSON.stringify(buildTerseManifest(commands, 'ask-marcel-office-cli', '0.0.0', stamp)).length;
    const terseDrive = JSON.stringify(filterManifestByCategory(buildTerseManifest(commands, 'ask-marcel-office-cli', '0.0.0', stamp), 'drive')).length;
    // If a band trips, update BOTH the documented "~N KB" hints and these bounds in the same change.
    expect(full).toBeGreaterThan(350_000); // floor: catch a broken/empty projection
    expect(full).toBeLessThan(520_000); // ceiling: docs say ~440 KB
    expect(terseAll).toBeGreaterThan(22_000);
    expect(terseAll).toBeLessThan(40_000); // docs say ~31 KB
    expect(terseDrive).toBeLessThan(8_500); // docs say ~6 KB/category (drive is the largest)
  });
});

describe('filterManifestByCategory — single-category projection', () => {
  const registry: Readonly<Record<string, Command>> = {
    'list-foo': fakeCmd({ category: 'mail' }),
    'list-bar': fakeCmd({ category: 'drive' }),
    'list-baz': fakeCmd({ category: 'mail' }),
  };

  it('keeps only commands whose category matches and preserves package/version/generatedAt', () => {
    const full = buildManifest(registry, 'fake-pkg', '0.0.1');
    const result = filterManifestByCategory(full, 'mail');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.package).toBe('fake-pkg');
    expect(result.value.commands.map((c) => c.name).toSorted((a, b) => a.localeCompare(b))).toEqual(['list-baz', 'list-foo']);
  });

  it('returns ok with empty commands when the requested category is valid but no commands match', () => {
    const full = buildManifest(registry, 'fake-pkg', '0.0.1');
    const result = filterManifestByCategory(full, 'excel');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.commands).toEqual([]);
  });

  it('rejects an unknown category with a discriminated error listing the available categories so the LLM can recover', () => {
    const full = buildManifest(registry, 'fake-pkg', '0.0.1');
    const result = filterManifestByCategory(full, 'notarealcategory');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'unknown_category') {
      expect(result.error.category).toBe('notarealcategory');
      expect(result.error.available).toContain('mail');
      expect(result.error.available).toContain('drive');
      expect(result.error.available).toContain('lifecycle');
    }
  });

  it('composes with --terse: a terse manifest filtered by category yields terse entries only in that category', () => {
    const terse = buildTerseManifest(registry, 'fake-pkg', '0.0.1');
    const result = filterManifestByCategory(terse, 'mail');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.commands.map((c) => c.name).toSorted((a, b) => a.localeCompare(b))).toEqual(['list-baz', 'list-foo']);
    // Terse-only fields: each command should have exactly name/summary/category.
    for (const c of result.value.commands) {
      expect(Object.keys(c).toSorted((a, b) => a.localeCompare(b))).toEqual(['category', 'name', 'summary']);
    }
  });
});

describe('renderSingleCommand', () => {
  it('returns Markdown for an existing registry command', () => {
    const registry: Readonly<Record<string, Command>> = { 'get-current-user': fakeCmd({ summary: 'returns the user' }) };
    const result = renderSingleCommand(registry, 'get-current-user');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('# `get-current-user`');
      expect(result.value).toContain('returns the user');
    }
  });

  it('returns Markdown for a lifecycle command (login/logout/update/docs/help-json) even when the registry is empty', () => {
    const result = renderSingleCommand({}, 'login');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('# `login`');
      expect(result.value).toContain('Authenticate against Microsoft Graph');
    }
  });

  it('returns unknown_command with the alphabetically merged registry+lifecycle list when the command is missing', () => {
    const registry: Readonly<Record<string, Command>> = { 'list-zebra': fakeCmd(), 'list-apple': fakeCmd() };
    const result = renderSingleCommand(registry, 'list-banana');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'unknown_command') {
      expect(result.error.name).toBe('list-banana');
      expect(result.error.available).toEqual(['docs', 'help-json', 'list-apple', 'list-zebra', 'login', 'logout', 'update']);
    }
  });
});
