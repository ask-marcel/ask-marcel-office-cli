import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleDocx, buildTrackedChangesDocx } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';

type Diff = { readonly contentType: string; readonly size: number; readonly text: string; readonly added?: number; readonly removed?: number; readonly note?: string };

const WEEK_6 = '# Plan\n\n- Kick-off 1 Oct\n- Budget 10k';
const WEEK_7 = '# Plan\n\n- Kick-off 8 Oct\n- Budget 10k\n- Owner: Robin Chen';

const drive = (files: Record<string, { readonly name?: string; readonly folder?: object | null; readonly text?: string }>): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    get: async (path: string): Promise<Result<unknown, GraphError>> => {
      const file = files[path];
      return file === undefined ? err({ type: 'api_error', status: 404, message: 'itemNotFound: gone', code: 'itemNotFound' }) : ok({ name: file.name, folder: file.folder });
    },
    getBinary: async (path: string): Promise<Result<unknown, GraphError>> => {
      const file = files[path.replace(/\/content$/, '')];
      return ok({ contentType: 'text/plain', size: file?.text?.length ?? 0, text: file?.text ?? '' });
    },
  });

const run = async (graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands['diff-drive-items'];
  if (!command) throw new Error('diff-drive-items is not registered');
  return command.execute(graph, params);
};

const PARAMS = { driveId: 'd1', itemId: 'i6', otherDriveId: 'd1', otherItemId: 'i7' };

describe('diff-drive-items', () => {
  it('answers the changed lines between two files as a unified diff, with the counts', async () => {
    const graph = drive({ '/drives/d1/items/i6': { name: 'update-6.md', text: WEEK_6 }, '/drives/d1/items/i7': { name: 'update-7.md', text: WEEK_7 } });
    const result = await run(graph, PARAMS);
    if (!result.ok) throw new Error(result.error.message);
    const diff = result.value as Diff;
    expect(diff.text).toBe('--- a/update-6.md\n+++ b/update-7.md\n@@ -1,4 +1,5 @@\n # Plan\n \n-- Kick-off 1 Oct\n+- Kick-off 8 Oct\n - Budget 10k\n+- Owner: Robin Chen\n');
    expect(diff).toMatchObject({ contentType: 'text/x-diff', added: 2, removed: 1, size: new TextEncoder().encode(diff.text).byteLength });
    expect(diff.note).toBeUndefined();
  });

  it('says so when the two files render the same', async () => {
    const graph = drive({ '/drives/d1/items/i6': { name: 'update-6.md', text: WEEK_6 }, '/drives/d1/items/i7': { name: 'update-7.md', text: WEEK_6 } });
    const result = await run(graph, PARAMS);
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toEqual({ contentType: 'text/x-diff', size: 0, text: '', added: 0, removed: 0, note: 'The two files render to the same markdown.' });
  });

  it('refuses a diff too long to be useful instead of printing both files', async () => {
    const lines = (prefix: string): string => Array.from({ length: 600 }, (_unused, i) => `${prefix} ${i}`).join('\n');
    const graph = drive({ '/drives/d1/items/i6': { name: 'a.md', text: lines('old') }, '/drives/d1/items/i7': { name: 'b.md', text: lines('new') } });
    const result = await run(graph, PARAMS);
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toEqual({
      contentType: 'text/x-diff',
      size: 0,
      text: '',
      note: 'The two renders differ by more than 1,000 added or removed lines (a changed line counts once each way), too many for a useful diff: read each file with download-drive-item-as-markdown.',
    });
  });

  it('passes on a failed read, names a folder, and refuses a file that has no text', async () => {
    const missing = await run(drive({ '/drives/d1/items/i6': { name: 'update-6.md', text: WEEK_6 } }), PARAMS);
    expect(missing.ok).toBe(false);
    if (!missing.ok) expect(missing.error.message).toBe('itemNotFound: gone');
    const folder = await run(drive({ '/drives/d1/items/i6': { name: 'Reports', folder: {} }, '/drives/d1/items/i7': { name: 'update-7.md', text: WEEK_7 } }), PARAMS);
    expect(folder.ok).toBe(false);
    if (!folder.ok) expect(folder.error).toEqual({ type: 'validation_error', message: 'Reports is a folder, not a file: pick a file inside it with list-folder-files.' });
    const image = await run(drive({ '/drives/d1/items/i6': { name: 'chart.png', text: 'x' }, '/drives/d1/items/i7': { name: 'update-7.md', text: WEEK_7 } }), PARAMS);
    expect(image.ok).toBe(false);
    if (!image.ok) expect(image.error.message).toContain('png');
  });
  it('refuses a call without the second file, and takes a null folder facet or a missing name as a plain file', async () => {
    const refused = await run(drive({}), { driveId: 'd1', itemId: 'i6', otherDriveId: 'd1' });
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.error.type).toBe('validation_error');
    const graph = drive({ '/drives/d1/items/i6': { folder: null, text: WEEK_6 }, '/drives/d1/items/i7': { name: 'update-7.md', text: WEEK_7 } });
    const result = await run(graph, PARAMS);
    if (!result.ok) throw new Error(result.error.message);
    expect((result.value as Diff).text).toStartWith('--- a/\n+++ b/update-7.md\n');
  });

  it('diffs the comments and tracked changes too with --include-metadata true', async () => {
    const plain = await buildSampleDocx();
    const tracked = await buildTrackedChangesDocx();
    const graph = fakeGraphClient({
      get: async (path: string) => ok({ name: path.endsWith('i6') ? 'week-6.docx' : 'week-7.docx' }),
      getBinary: async (path: string) => {
        const bytes = path.includes('/i6/') ? plain : tracked;
        return ok({ contentType: 'application/octet-stream', size: bytes.byteLength, base64: Buffer.from(bytes).toString('base64') });
      },
    });
    const withMeta = await run(graph, { ...PARAMS, includeMetadata: 'true' });
    const without = await run(graph, { ...PARAMS, includeMetadata: 'false' });
    if (!withMeta.ok || !without.ok) throw new Error('expected ok');
    expect((withMeta.value as Diff).text).toContain('Robin Chen');
    expect((without.value as Diff).text).not.toContain('Tracked changes');
  });
});
