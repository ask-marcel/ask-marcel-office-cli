import { describe, expect, it } from 'bun:test';
import { err, ok, type Result } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './download-drive-item-version.ts';

const VERSIONS = '/drives/d1/items/i1/versions?$select=id,lastModifiedDateTime';
const listing = {
  value: [
    { id: '3.0', lastModifiedDateTime: '2026-09-18T09:00:00Z' },
    { id: '2.0', lastModifiedDateTime: '2026-09-17T09:00:00Z' },
    { id: '1.0', lastModifiedDateTime: '2026-09-16T09:00:00Z' },
    { id: 'x' },
  ],
};

const graphWith = (versions: Result<unknown, GraphError> = ok(listing)): { paths: string[]; downloads: string[]; graph: ReturnType<typeof fakeGraphClient> } => {
  const paths: string[] = [];
  const downloads: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      if (path === VERSIONS) return versions;
      return ok({ name: 'plan.docx' });
    },
    getBinaryElevated: async (path) => {
      downloads.push(path);
      return ok({ contentType: 'application/octet-stream', size: 3, base64: 'AAAA' });
    },
  });
  return { paths, downloads, graph };
};

describe('picking a historical version by date', () => {
  it('downloads the newest version saved strictly before the instant and says which one it chose', async () => {
    const { paths, downloads, graph } = graphWith();
    const result = await execute(graph, { driveId: 'd1', itemId: 'i1', before: '2026-09-17T09:00:00Z' });
    expect(paths).toEqual([VERSIONS]);
    expect(downloads).toEqual(['/drives/d1/items/i1/versions/1.0/content']);
    if (!result.ok) throw new Error('expected ok');
    expect(result.value).toEqual({ contentType: 'application/octet-stream', size: 3, base64: 'AAAA', versionId: '1.0' });
  });

  it('accepts the relative date vocabulary and lets an explicit --version-id win without listing, still reporting the id used', async () => {
    const relative = graphWith();
    const r1 = await execute(relative.graph, { driveId: 'd1', itemId: 'i1', before: 'now' });
    expect(r1.ok).toBe(true);
    expect(relative.downloads).toEqual(['/drives/d1/items/i1/versions/3.0/content']);
    const explicit = graphWith();
    const r2 = await execute(explicit.graph, { driveId: 'd1', itemId: 'i1', versionId: '2', before: '2026-09-17T09:00:00Z' });
    if (!r2.ok) throw new Error('expected ok');
    expect(explicit.paths).toEqual([]);
    expect(explicit.downloads).toEqual(['/drives/d1/items/i1/versions/2.0/content']);
    expect((r2.value as Record<string, unknown>)['versionId']).toBe('2.0');
  });

  it('answers a clear not-found when nothing was saved before the instant', async () => {
    const { downloads, graph } = graphWith();
    const result = await execute(graph, { driveId: 'd1', itemId: 'i1', before: '2026-09-16T09:00:00Z' });
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.status).toBe(404);
    expect(result.error.code).toBe('cli_no_version_before');
    expect(result.error.message).toContain('before 2026-09-16T09:00:00Z');
    expect(downloads).toEqual([]);
  });

  it('refuses a call with neither flag, an unreadable date, and passes a failed listing through', async () => {
    const neither = await execute(graphWith().graph, { driveId: 'd1', itemId: 'i1' });
    expect(neither.ok).toBe(false);
    if (!neither.ok) {
      expect(neither.error.type).toBe('validation_error');
      expect(neither.error.message).toContain('--before');
    }
    const unreadable = await execute(graphWith().graph, { driveId: 'd1', itemId: 'i1', before: 'whenever' });
    expect(unreadable.ok).toBe(false);
    if (!unreadable.ok) expect(unreadable.error.type).toBe('validation_error');
    const failing = graphWith(err({ type: 'api_error', status: 403, message: 'no' }));
    const failed = await execute(failing.graph, { driveId: 'd1', itemId: 'i1', before: 'now' });
    expect(failed.ok).toBe(false);
    if (!failed.ok) expect(failed.error.message).toBe('no');
  });
});
