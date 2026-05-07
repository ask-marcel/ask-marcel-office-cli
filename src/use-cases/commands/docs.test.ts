import { describe, expect, it } from 'bun:test';
import type { Command } from './command-types.ts';
import { buildManifest, renderSingleCommand } from './docs.ts';

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
    example: 'ask-marcel fake',
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

  it('uses the real `new Date()` when no clock injector is given', () => {
    const before = Date.now();
    const manifest = buildManifest({ foo: fakeCmd() }, 'fake-pkg', '0.0.1');
    const after = Date.now();
    const generatedAt = new Date(manifest.generatedAt).getTime();
    expect(generatedAt).toBeGreaterThanOrEqual(before);
    expect(generatedAt).toBeLessThanOrEqual(after);
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
