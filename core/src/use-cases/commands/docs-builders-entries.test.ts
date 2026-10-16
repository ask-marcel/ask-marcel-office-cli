import { describe, expect, it } from 'bun:test';
import { fakeCommand, fakeCommandPackage } from '../../test-helpers/command-package-fake.ts';
import { buildManifest, buildTerseManifest, filterManifestByCategory } from './docs-builders.ts';

/*
 * What the manifest says about one command (core, package split phase 3): the
 * meta it always carries, each optional field only when the command sets it,
 * and the cursor kind of a paginated command.
 */

const AT = (): Date => new Date('2026-10-15T00:00:00.000Z');

const entryOf = (meta: Parameters<typeof fakeCommand>[0]): unknown =>
  buildManifest(fakeCommandPackage({ registry: { 'list-widgets': fakeCommand(meta) }, lifecycle: [] }), '1.0.0', AT).commands[0];

const BASE = {
  name: 'list-widgets',
  summary: 'Do a widget thing. Second sentence.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/widgets',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/widgets',
  options: [{ name: 'top', key: 'top', required: false, description: 'How many widgets.' }],
  example: 'widget-cli list-widgets --top 3',
  effect: 'read',
};

describe('the manifest entry of a command', () => {
  it('carries only the meta every command has when the command sets nothing else', () => {
    expect(entryOf({})).toEqual(BASE);
  });

  it('carries each optional field the command sets', () => {
    const full = {
      positionalArguments: [{ name: 'path', required: true, description: 'A file.' }],
      responseShape: '{ id }',
      bodyTemplate: '{ "requests": [] }',
      pagination: true,
      paginationStrategy: 'deltaLink',
      scopesRequired: ['Files.Read'],
      needsElevatedToken: true,
      needsSubstrateToken: 'ic3',
      producesBytes: true,
      producesMedia: true,
      stability: 'experimental',
    } as const;
    expect(entryOf(full)).toEqual({ ...BASE, ...full });
  });

  it('gives a paginated command the nextLink cursor when it names none', () => {
    expect(entryOf({ pagination: true })).toEqual({ ...BASE, pagination: true, paginationStrategy: 'nextLink' });
  });

  it('leaves out an empty scope list', () => {
    expect(entryOf({ scopesRequired: [] })).toEqual(BASE);
  });

  it('carries the stability of a command in the terse view too', () => {
    const source = fakeCommandPackage({ registry: { 'list-widgets': fakeCommand({ stability: 'experimental' }), 'get-widget': fakeCommand() }, lifecycle: [] });
    expect(buildTerseManifest(source, '1.0.0', AT).commands).toEqual([
      { name: 'get-widget', summary: 'Do a widget thing.', category: 'drive' },
      { name: 'list-widgets', summary: 'Do a widget thing.', category: 'drive', stability: 'experimental' },
    ]);
  });
});

describe('a manifest filtered by category', () => {
  it('keeps the commands of the category, and refuses an unknown one with the categories sorted', () => {
    const manifest = buildManifest(fakeCommandPackage(), '1.0.0', AT);
    const filtered = filterManifestByCategory(manifest, 'meta');
    expect(filtered.ok && filtered.value.commands.map((c) => c.name)).toEqual(['find-widgets']);
    expect(filterManifestByCategory(manifest, 'widgets')).toEqual({
      ok: false,
      error: {
        type: 'unknown_category',
        category: 'widgets',
        available: ['calendar', 'chats', 'drive', 'excel', 'lifecycle', 'mail', 'meta', 'notes', 'sharepoint', 'tasks', 'teams', 'user'],
      },
    });
  });
});
