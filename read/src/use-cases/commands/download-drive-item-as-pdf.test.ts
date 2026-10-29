import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { execute } from './download-drive-item-as-pdf.ts';

const TENANT = '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04';
const PARAMS = { driveId: 'd1', itemId: 'i1' };
const FAILED: GraphError = { type: 'api_error', status: 503, message: 'busy' };

// A drive item named `item`, whose bytes are `bytes`, and a record of each call by its route.
const graphFor = (item: unknown, bytes: Result<unknown, GraphError>): { readonly graph: GraphClient; readonly calls: string[] } => {
  const calls: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      calls.push(`get ${path}`);
      return ok(item);
    },
    getGuest: async (path, tenant) => {
      calls.push(`getGuest ${path} ${tenant}`);
      return ok(item);
    },
    getBinary: async (path) => {
      calls.push(`getBinary ${path}`);
      return bytes;
    },
    getBinaryGuest: async (path, tenant) => {
      calls.push(`getBinaryGuest ${path} ${tenant}`);
      return bytes;
    },
  });
  return { graph, calls };
};

describe('downloading a drive item as PDF', () => {
  it('refuses a folder, naming it and the command that lists its children, before any byte fetch', async () => {
    const { graph, calls } = graphFor({ name: 'Plans', folder: { childCount: 2 } }, ok({}));
    expect(await execute(graph, PARAMS)).toEqual(
      err({
        type: 'api_error',
        status: 400,
        message: "item 'Plans' is a folder, not a file — use `list-folder-files --drive-id d1 --item-id i1` to enumerate its children, then pick a file to convert.",
      })
    );
    expect(calls).toEqual(['get /drives/d1/items/i1']);
  });

  it('names a nameless folder as the empty name', async () => {
    const { graph } = graphFor({ folder: {} }, ok({}));
    const result = await execute(graph, PARAMS);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message.startsWith("item '' is a folder")).toBe(true);
  });

  it('treats a null folder facet as a file', async () => {
    const { graph, calls } = graphFor({ name: 'notes.txt', folder: null }, ok({ base64: btoa('hi') }));
    expect((await execute(graph, PARAMS)).ok).toBe(true);
    expect(calls).toEqual(['get /drives/d1/items/i1', 'getBinary /drives/d1/items/i1/content']);
  });

  it('returns a plain-text source as text, without asking Graph to convert it', async () => {
    const { graph, calls } = graphFor({ name: 'notes.txt' }, ok({ base64: btoa('hello') }));
    expect(await execute(graph, PARAMS)).toEqual(
      ok({
        contentType: 'text/plain',
        size: 5,
        text: 'hello',
        passthrough: true,
        note: 'source is plain-text (notes.txt); raw text returned without Graph format=pdf conversion',
      })
    );
    expect(calls).toEqual(['get /drives/d1/items/i1', 'getBinary /drives/d1/items/i1/content']);
  });

  it('reports the byte fetch failure of a plain-text source', async () => {
    const { graph } = graphFor({ name: 'notes.txt' }, err(FAILED));
    expect(await execute(graph, PARAMS)).toEqual(err(FAILED));
  });

  it('returns a PDF source as its own bytes, without asking Graph to convert it', async () => {
    const { graph, calls } = graphFor({ name: 'deck.pdf' }, ok({ contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' }));
    expect(await execute(graph, PARAMS)).toEqual(
      ok({
        contentType: 'application/pdf',
        size: 5,
        base64: 'JVBERi0=',
        passthrough: true,
        note: 'source is already PDF (deck.pdf); raw bytes returned without Graph format=pdf conversion',
      })
    );
    expect(calls).toEqual(['get /drives/d1/items/i1', 'getBinary /drives/d1/items/i1/content']);
  });

  it('reports the byte fetch failure of a PDF source', async () => {
    const { graph } = graphFor({ name: 'deck.pdf' }, err(FAILED));
    expect(await execute(graph, PARAMS)).toEqual(err(FAILED));
  });

  it('asks Graph to convert any other source', async () => {
    const { graph, calls } = graphFor({ name: 'plan.docx' }, ok({ contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' }));
    expect((await execute(graph, PARAMS)).ok).toBe(true);
    expect(calls).toEqual(['get /drives/d1/items/i1', 'getBinary /drives/d1/items/i1/content?format=pdf']);
  });

  it('reads the metadata and the bytes of a partner-tenant file with the guest token', async () => {
    for (const name of ['notes.txt', 'deck.pdf']) {
      const { graph, calls } = graphFor({ name }, ok({ contentType: 'application/pdf', size: 2, base64: btoa('hi') }));
      expect((await execute(graph, { ...PARAMS, tenantId: TENANT })).ok).toBe(true);
      expect(calls).toEqual([`getGuest /drives/d1/items/i1 ${TENANT}`, `getBinaryGuest /drives/d1/items/i1/content ${TENANT}`]);
    }
  });

  it('refuses a malformed --tenant-id before any call', async () => {
    const { graph, calls } = graphFor({ name: 'notes.txt' }, ok({}));
    const result = await execute(graph, { ...PARAMS, tenantId: 'contoso' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toEqual(expect.objectContaining({ type: 'validation_error', code: 'invalid_tenant_id' }));
    expect(calls).toEqual([]);
  });
});
