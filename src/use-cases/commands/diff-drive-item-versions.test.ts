import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleDocx, buildTrackedChangesDocx } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';

const VERSIONS = '/drives/d1/items/i1/versions?$select=id,lastModifiedDateTime';
const saved = (id: string, day: number): { id: string; lastModifiedDateTime: string } => ({ id, lastModifiedDateTime: `2026-09-${day}T09:00:00Z` });
const LISTING = { value: [saved('6.0', 23), saved('5.0', 22), saved('4.0', 21)] };

type Files = { readonly name: string; readonly versions: Record<string, string>; readonly current: string };

const graphFor = (files: Files, downloads: string[] = []): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    get: async (path: string): Promise<Result<unknown, GraphError>> => ok(path === VERSIONS ? LISTING : { name: files.name }),
    getBinaryElevated: async (path: string): Promise<Result<unknown, GraphError>> => {
      downloads.push(path);
      const text = files.versions[path] ?? '';
      return ok({ contentType: 'text/plain', size: text.length, text });
    },
    getBinary: async (path: string): Promise<Result<unknown, GraphError>> => {
      downloads.push(path);
      return ok({ contentType: 'text/plain', size: files.current.length, text: files.current });
    },
  });

const run = async (graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands['diff-drive-item-versions'];
  if (!command) throw new Error('diff-drive-item-versions is not registered');
  return command.execute(graph, { driveId: 'd1', itemId: 'i1', ...params });
};

const HANDOVER: Files = {
  name: 'handover.md',
  versions: { '/drives/d1/items/i1/versions/5.0/content': '- Keys: reception\n- Laptop: returned', '/drives/d1/items/i1/versions/4.0/content': '- Keys: reception' },
  current: '- Keys: reception\n- Laptop: returned\n- Badge: returned',
};

describe('diff-drive-item-versions', () => {
  it('compares the version saved before an instant with the live file, and says which version it chose', async () => {
    const downloads: string[] = [];
    const result = await run(graphFor(HANDOVER, downloads), { before: '2026-09-22T12:00:00Z' });
    if (!result.ok) throw new Error(result.error.message);
    const text = '--- a/handover.md (version 5.0)\n+++ b/handover.md (current)\n@@ -1,2 +1,3 @@\n - Keys: reception\n - Laptop: returned\n+- Badge: returned\n';
    expect(result.value).toEqual({
      contentType: 'text/x-diff',
      size: new TextEncoder().encode(text).byteLength,
      text,
      added: 1,
      removed: 0,
      versionId: '5.0',
    });
    expect(downloads).toEqual(['/drives/d1/items/i1/versions/5.0/content', '/drives/d1/items/i1/content']);
  });

  it('takes an explicit version id, integer or not', async () => {
    const result = await run(graphFor(HANDOVER), { versionId: '4' });
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toMatchObject({ added: 2, removed: 0, versionId: '4.0' });
    const dotted = await run(graphFor(HANDOVER), { versionId: '5.0', includeMetadata: 'false' });
    if (!dotted.ok) throw new Error(dotted.error.message);
    expect(dotted.value).toMatchObject({ added: 1, removed: 0, versionId: '5.0' });
  });

  it('refuses a date it cannot read, and passes on a failed read of the file, of the version or of the live file', async () => {
    const bad = await run(graphFor(HANDOVER), { before: 'someday' });
    expect(bad.ok).toBe(false);
    if (!bad.ok) expect(bad.error.type).toBe('validation_error');
    const gone = { type: 'api_error' as const, status: 404, message: 'itemNotFound: gone' };
    const failing = (where: 'meta' | 'version' | 'live'): ReturnType<typeof fakeGraphClient> =>
      fakeGraphClient({
        get: async (path: string) => {
          if (path === VERSIONS) return ok(LISTING);
          return where === 'meta' ? err(gone) : ok({ name: 'handover.md' });
        },
        getBinaryElevated: async () => (where === 'version' ? err(gone) : ok({ contentType: 'text/plain', size: 1, text: 'a' })),
        getBinary: async () => (where === 'live' ? err(gone) : ok({ contentType: 'text/plain', size: 1, text: 'b' })),
      });
    for (const where of ['meta', 'version', 'live'] as const) {
      const result = await run(failing(where), { before: '2026-09-22T12:00:00Z' });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.message).toContain('itemNotFound: gone');
    }
  });

  it('answers with no diff when nothing was saved after the instant, without downloading a thing', async () => {
    const downloads: string[] = [];
    const result = await run(graphFor(HANDOVER, downloads), { before: '2026-09-24T00:00:00Z' });
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toEqual({
      contentType: 'text/x-diff',
      size: 0,
      text: '',
      added: 0,
      removed: 0,
      versionId: '6.0',
      note: 'Nothing was saved after 2026-09-24T00:00:00Z: version 6.0 is the live file.',
    });
    expect(downloads).toEqual([]);
  });

  it('says so when the version and the live file render the same', async () => {
    const same: Files = { ...HANDOVER, current: '- Keys: reception\n- Laptop: returned' };
    const result = await run(graphFor(same), { before: '2026-09-22T12:00:00Z' });
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toMatchObject({ text: '', note: 'This version and the live file render to the same markdown.' });
  });

  it('refuses a Loop page, whose old versions Graph cannot render', async () => {
    const loop = await run(graphFor({ ...HANDOVER, name: 'notes.loop' }), { before: '2026-09-22T12:00:00Z' });
    expect(loop.ok).toBe(false);
    if (!loop.ok) expect(loop.error).toMatchObject({ status: 415, code: 'unsupported_version_render' });
  });

  it('diffs the comments and tracked changes too with --include-metadata true', async () => {
    const plain = await buildSampleDocx();
    const tracked = await buildTrackedChangesDocx();
    const graph = fakeGraphClient({
      get: async (path: string) => ok(path === VERSIONS ? LISTING : { name: 'handover.docx' }),
      getBinaryElevated: async () => ok({ contentType: 'application/octet-stream', size: plain.byteLength, base64: Buffer.from(plain).toString('base64') }),
      getBinary: async () => ok({ contentType: 'application/octet-stream', size: tracked.byteLength, base64: Buffer.from(tracked).toString('base64') }),
    });
    const withMeta = await run(graph, { before: '2026-09-22T12:00:00Z', includeMetadata: 'true' });
    const without = await run(graph, { before: '2026-09-22T12:00:00Z' });
    if (!withMeta.ok || !without.ok) throw new Error('expected ok');
    expect((withMeta.value as { text: string }).text).toContain('Robin Chen');
    expect((without.value as { text: string }).text).not.toContain('Tracked changes');
  });
});
