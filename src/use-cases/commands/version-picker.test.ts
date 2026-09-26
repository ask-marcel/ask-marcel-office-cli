import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './download-drive-item-version.ts';

const VERSIONS = '/drives/d1/items/i1/versions?$select=id,lastModifiedDateTime';

describe('the version saved before an instant', () => {
  it('is the newest dated one, whatever order Graph lists them in, skipping entries without an id or a date', async () => {
    const downloads: string[] = [];
    const listing = {
      value: [
        { id: '9.0' },
        { id: '4.0', lastModifiedDateTime: '2026-09-21T09:00:00Z' },
        { lastModifiedDateTime: '2026-09-22T11:00:00Z' },
        { id: '6.0', lastModifiedDateTime: '2026-09-23T09:00:00Z' },
        { id: '5.0', lastModifiedDateTime: '2026-09-22T09:00:00Z' },
      ],
    };
    const graph = fakeGraphClient({
      get: async (path: string) => ok(path === VERSIONS ? listing : { name: 'plan.docx' }),
      getBinaryElevated: async (path: string) => {
        downloads.push(path);
        return ok({ contentType: 'application/octet-stream', size: 3, base64: 'AAAA' });
      },
    });
    const result = await execute(graph, { driveId: 'd1', itemId: 'i1', before: '2026-09-22T12:00:00Z' });
    if (!result.ok) throw new Error(result.error.message);
    expect(downloads).toEqual(['/drives/d1/items/i1/versions/5.0/content']);
    expect(result.value).toMatchObject({ versionId: '5.0' });
  });
});
