import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const VERSIONS = '/drives/d1/items/i1/versions?$select=id,lastModifiedDateTime';
const saved = (id: string, day: number): { id: string; lastModifiedDateTime: string } => ({ id, lastModifiedDateTime: `2026-09-${day}T09:00:00Z` });

type Recorded = { readonly graph: ReturnType<typeof fakeGraphClient>; readonly elevated: string[]; readonly basic: string[] };

// Version history listed newest first as Graph would, but deliberately out of order to prove the sort.
const history = (listing: ReadonlyArray<{ id: string; lastModifiedDateTime: string }>, name = 'handover.md'): Recorded => {
  const elevated: string[] = [];
  const basic: string[] = [];
  const graph = fakeGraphClient({
    get: async (path: string): Promise<Result<unknown, GraphError>> => ok(path === VERSIONS ? { value: listing } : { name }),
    getBinaryElevated: async (path: string): Promise<Result<unknown, GraphError>> => {
      elevated.push(path);
      const text = path.includes('/versions/') ? '- Keys: reception' : '- Keys: reception\n- Badge: returned';
      return ok({ contentType: 'text/plain', size: text.length, text });
    },
    getBinary: async (path: string): Promise<Result<unknown, GraphError>> => {
      basic.push(path);
      const text = '- Keys: reception\n- Badge: returned';
      return ok({ contentType: 'text/plain', size: text.length, text });
    },
  });
  return { graph, elevated, basic };
};

const run = async (name: string, graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  return command.execute(graph, { driveId: 'd1', itemId: 'i1', ...params });
};

describe('which version a version command reads when none is named', () => {
  it('diffs what the last save changed when neither --version-id nor --before is given', async () => {
    const h = history([saved('5.0', 22), saved('6.0', 23), saved('4.0', 21)]);
    const result = await run('diff-drive-item-versions', h.graph, {});
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toMatchObject({ added: 1, removed: 0, versionId: '5.0' });
    expect(h.elevated).toEqual(['/drives/d1/items/i1/versions/5.0/content']);
    expect(h.basic).toEqual(['/drives/d1/items/i1/content']);
  });

  it('says a file saved only once has nothing earlier to compare, without downloading it', async () => {
    const h = history([saved('1.0', 20)]);
    const result = await run('diff-drive-item-versions', h.graph, {});
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toEqual({
      contentType: 'text/x-diff',
      size: 0,
      text: '',
      added: 0,
      removed: 0,
      versionId: '1.0',
      note: 'Version 1.0 is the only one saved: there is no earlier version to compare with the live file.',
    });
    expect([...h.elevated, ...h.basic]).toEqual([]);
  });

  it('reads the live file when --before lands after the last save, since Graph refuses the live version as a version', async () => {
    const h = history([saved('3.0', 18), saved('2.0', 17)]);
    const result = await run('download-drive-item-version', h.graph, { before: '2026-09-19T00:00:00Z' });
    if (!result.ok) throw new Error(result.error.message);
    expect(h.elevated).toEqual(['/drives/d1/items/i1/content']);
    expect(result.value).toMatchObject({ versionId: '3.0', current: true, note: 'Nothing was saved after 2026-09-19T00:00:00Z: this is the live file.' });
  });

  it('renders the live file as markdown or PDF the same way when --before lands after the last save', async () => {
    const markdown = history([saved('3.0', 18)], 'handover.md');
    const md = await run('download-drive-item-version', markdown.graph, { before: '2026-09-19T00:00:00Z', format: 'markdown' });
    if (!md.ok) throw new Error(md.error.message);
    expect(markdown.elevated).toEqual(['/drives/d1/items/i1/content']);
    expect(md.value).toMatchObject({ text: '- Keys: reception\n- Badge: returned', current: true });
    const pdf = history([saved('3.0', 18)], 'plan.docx');
    const rendered = await run('download-drive-item-version', pdf.graph, { before: '2026-09-19T00:00:00Z', format: 'pdf' });
    expect(rendered.ok).toBe(true);
    expect(pdf.elevated).toEqual(['/drives/d1/items/i1/content?format=pdf']);
  });

  it('answers a clear not-found when Graph lists no version at all', async () => {
    const result = await run('diff-drive-item-versions', history([]).graph, {});
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toMatchObject({ type: 'api_error', status: 404, code: 'cli_no_versions' });
  });

  it('passes a failed version listing through when neither flag is given, and names the listing when it is empty', async () => {
    const refused = fakeGraphClient({
      get: async (path: string) => (path === VERSIONS ? err({ type: 'api_error', status: 403, message: 'accessDenied: no', code: 'accessDenied' }) : ok({ name: 'handover.md' })),
    });
    const failed = await run('diff-drive-item-versions', refused, {});
    expect(failed.ok).toBe(false);
    if (!failed.ok) expect(failed.error).toMatchObject({ code: 'accessDenied' });
    const none = await run('diff-drive-item-versions', history([]).graph, {});
    if (!none.ok) expect(none.error.message).toContain('Graph listed no version of this file; list-drive-item-versions shows what exists');
  });

  it('keeps a PDF passthrough note beside the live-file note, and reads a live PDF as-is', async () => {
    const h = history([saved('3.0', 18)], 'contract.pdf');
    const result = await run('download-drive-item-version', h.graph, { before: '2026-09-19T00:00:00Z', format: 'pdf' });
    if (!result.ok) throw new Error(result.error.message);
    expect(h.elevated).toEqual(['/drives/d1/items/i1/content']);
    expect((result.value as { note: string }).note).toBe(
      'source is already PDF (contract.pdf); raw bytes returned without Graph format=pdf conversion Nothing was saved after 2026-09-19T00:00:00Z: this is the live file.'
    );
  });

  it('says which version Graph left unconverted when --format pdf answers the source bytes', async () => {
    const h = history([saved('3.0', 18), saved('2.0', 17)], 'plan.docx');
    const result = await run('download-drive-item-version', h.graph, { versionId: '2.0', format: 'pdf', includeMetadata: 'false' });
    if (!result.ok) throw new Error(result.error.message);
    expect((result.value as { note: string }).note).toContain('for version 2.0 of plan.docx');
    const original = await run('download-drive-item-version', h.graph, { versionId: '2.0', format: 'original' });
    expect(original.ok).toBe(true);
  });

  it('passes a failed read of a PDF version through', async () => {
    const failing = fakeGraphClient({
      get: async () => ok({ name: 'contract.pdf' }),
      getBinaryElevated: async () => err({ type: 'api_error', status: 404, message: 'itemNotFound: gone', code: 'itemNotFound' }),
    });
    const result = await run('download-drive-item-version', failing, { versionId: '2.0', format: 'pdf' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toMatchObject({ code: 'itemNotFound' });
  });

  it('keeps reading a real historical version from its version path', async () => {
    const h = history([saved('3.0', 18), saved('2.0', 17)]);
    const result = await run('download-drive-item-version', h.graph, { before: '2026-09-18T00:00:00Z' });
    if (!result.ok) throw new Error(result.error.message);
    expect(h.elevated).toEqual(['/drives/d1/items/i1/versions/2.0/content']);
    expect(result.value).not.toHaveProperty('current');
    expect(result.value).toMatchObject({ versionId: '2.0' });
  });
});
