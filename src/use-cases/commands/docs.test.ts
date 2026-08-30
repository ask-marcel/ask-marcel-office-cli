import { describe, expect, it } from 'bun:test';
import type { Command } from './command-types.ts';
import type { CommandManifestEntry } from './docs-render.ts';
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

const LIFECYCLE_NAMES = ['docs', 'help-json', 'login', 'logout', 'mcp', 'update'] as const;

describe('buildManifest', () => {
  it('builds a manifest with package name, version, generatedAt, and registry+lifecycle commands sorted alphabetically', () => {
    const registry: Readonly<Record<string, Command>> = { 'list-zebra': fakeCmd(), 'list-apple': fakeCmd() };
    const manifest = buildManifest(registry, 'fake-pkg', '0.0.1', () => new Date('2026-04-30T12:00:00Z'));
    expect(manifest.package).toBe('fake-pkg');
    expect(manifest.version).toBe('0.0.1');
    expect(manifest.generatedAt).toBe('2026-04-30T12:00:00.000Z');
    expect(manifest.commands.map((c) => c.name)).toEqual(['docs', 'help-json', 'list-apple', 'list-zebra', 'login', 'logout', 'mcp', 'update']);
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

  // 2026-07-24: the manifest carries one name per command; commandAliases
  // serialization was removed with the alias system.
  it('never serializes a commandAliases key, since a command has exactly one name', () => {
    const manifest = buildManifest({ 'aaa-plain': fakeCmd() }, 'fake-pkg', '0.0.1');
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
      expect(result.error.available).toEqual(['docs', 'help-json', 'list-apple', 'list-zebra', 'login', 'logout', 'mcp', 'update']);
    }
  });
});

describe('buildManifest — the conditional fields toEntry adds or omits', () => {
  const entryFor = (name: string, meta: Partial<Command['meta']>): CommandManifestEntry | undefined =>
    buildManifest({ [name]: fakeCmd(meta) }, 'fake-pkg', '0.0.1').commands.find((c) => c.name === name);

  it('defaults a paginated command to the nextLink cursor strategy', () => {
    expect(entryFor('aaa-paged', { pagination: true })?.paginationStrategy).toBe('nextLink');
  });

  it('lets a command override the default cursor strategy', () => {
    expect(entryFor('aaa-noskip', { pagination: true, paginationStrategy: 'nextLinkNoSkip' })?.paginationStrategy).toBe('nextLinkNoSkip');
  });

  it('omits paginationStrategy on a command that does not paginate', () => {
    expect(entryFor('aaa-plain', {})).not.toHaveProperty('paginationStrategy');
  });

  it('prefers an inline scopesRequired over the central graph-scopes map', () => {
    // `aaa-inline` is absent from GRAPH_SCOPES_BY_COMMAND, so the central lookup
    // yields undefined: only the inline value can populate the field here.
    expect(entryFor('aaa-inline', { scopesRequired: ['Mail.Read'] })?.scopesRequired).toEqual(['Mail.Read']);
  });

  it('omits scopesRequired when the command declares an empty list', () => {
    // An empty array is truthy, so only the explicit length check keeps
    // `scopesRequired: []` out of the manifest.
    expect(entryFor('aaa-noscope', { scopesRequired: [] })).not.toHaveProperty('scopesRequired');
  });

  it('omits positionalArguments / bodyTemplate / needsElevatedToken / needsSubstrateToken when the meta leaves them unset', () => {
    const entry = entryFor('aaa-bare', {});
    expect(entry).not.toHaveProperty('positionalArguments');
    expect(entry).not.toHaveProperty('bodyTemplate');
    expect(entry).not.toHaveProperty('needsElevatedToken');
    expect(entry).not.toHaveProperty('needsSubstrateToken');
  });

  it('carries positionalArguments / bodyTemplate / needsElevatedToken / needsSubstrateToken through when the meta sets them', () => {
    const positionalArguments = [{ name: 'target', required: true, description: 'the thing to act on' }];
    const entry = entryFor('aaa-full', { positionalArguments, bodyTemplate: '{"a":1}', needsElevatedToken: true, needsSubstrateToken: 'ic3' });
    expect(entry?.positionalArguments).toEqual(positionalArguments);
    expect(entry?.bodyTemplate).toBe('{"a":1}');
    expect(entry?.needsElevatedToken).toBe(true);
    // `needsSubstrateToken` is emitted by toEntry and ships on 5 commands, but
    // CommandManifestEntry never declares it: the conditional spread that adds it
    // slips past excess-property checking. Asserted through a cast rather than by
    // widening the type, since editing docs-render.ts pulls a separate file into
    // the mutation gate. See the follow-up note in this suite's sibling task.
    expect((entry as { needsSubstrateToken?: string } | undefined)?.needsSubstrateToken).toBe('ic3');
  });
});

describe('lifecycle entries — the invariants the registry sweep never reaches', () => {
  // login/logout/update/docs/help-json/mcp have no command file, so they are absent
  // from the `commands` registry and skip every invariant meta.test.ts applies to
  // the other 186. buildManifest merging them in is the only public route to them.
  const lifecycle = buildManifest({}, 'fake-pkg', '0.0.1').commands;
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
      // The example is what an LLM copies verbatim, so it has to be a real invocation.
      expect(entry.example.startsWith(`ask-marcel-office ${name}`)).toBe(true);
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

  it('keeps `docs` taking its command name as a required positional rather than a flag', () => {
    const positionals = byName('docs')?.positionalArguments ?? [];
    expect(positionals.map((p) => p.name)).toEqual(['command']);
    expect(positionals[0]?.required).toBe(true);
    expect(byName('docs')?.options).toEqual([]);
  });
});

describe('docs error payloads — discriminants and ordering', () => {
  const registry: Readonly<Record<string, Command>> = { 'list-foo': fakeCmd({ category: 'mail' }) };

  it('names the unknown_category discriminant outside any narrowing guard', () => {
    const result = filterManifestByCategory(buildManifest(registry, 'fake-pkg', '0.0.1'), 'notarealcategory');
    expect(result.ok).toBe(false);
    // Asserted unconditionally on purpose: a guard of the form
    // `if (result.error.type === '...')` skips its own body when the discriminant
    // changes, so it can never pin the discriminant it tests.
    if (!result.ok) expect(result.error.type).toBe('unknown_category');
  });

  it('lists the available categories alphabetically, not in declaration order', () => {
    const result = filterManifestByCategory(buildManifest(registry, 'fake-pkg', '0.0.1'), 'notarealcategory');
    if (!result.ok && result.error.type === 'unknown_category') {
      const { available } = result.error;
      expect(available.length).toBeGreaterThan(1);
      expect(available).toEqual([...available].toSorted((a, b) => a.localeCompare(b)));
    }
  });

  it('names the unknown_command discriminant outside any narrowing guard', () => {
    const result = renderSingleCommand(registry, 'list-banana');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('unknown_command');
  });
});
