import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildCommentedXlsx } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';

const command = commands['list-document-comments'];
const ARGS = { driveId: 'b!x', itemId: '01ABC' };
const PARTNER = '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04';

const binaryOf = (bytes: Uint8Array): Result<unknown, GraphError> =>
  ok({ contentType: 'application/octet-stream', size: bytes.byteLength, base64: Buffer.from(bytes).toString('base64') });

describe('list-document-comments', () => {
  it('reads the item name, downloads the file and lists its comments', async () => {
    if (!command) throw new Error('list-document-comments is not registered');
    const bytes = await buildCommentedXlsx();
    const calls: string[] = [];
    const graph = fakeGraphClient({
      get: async (path) => {
        calls.push(`get ${path}`);
        return ok({ name: 'plan.xlsx' });
      },
      getBinary: async (path) => {
        calls.push(`bin ${path}`);
        return binaryOf(bytes);
      },
    });
    const r = await command.execute(graph, ARGS);
    if (!r.ok) throw new Error(r.error.message);
    expect(calls).toEqual(['get /drives/b!x/items/01ABC', 'bin /drives/b!x/items/01ABC/content']);
    expect(r.value).toMatchObject({ name: 'plan.xlsx', format: 'xlsx', count: 3 });
    expect((r.value as { comments: ReadonlyArray<{ anchor: string }> }).comments.map((c) => c.anchor)).toEqual(["'Q3 Plan'!B2", 'Summary!C3', 'Summary!C3']);
  });

  it('refuses a folder, pointing at list-folder-files, before downloading anything', async () => {
    if (!command) throw new Error('list-document-comments is not registered');
    const r = await command.execute(fakeGraphClient({ get: async () => ok({ name: 'Plans', folder: { childCount: 3 } }) }), ARGS);
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error).toMatchObject({ type: 'api_error', status: 400 });
      expect(r.error.message).toContain('list-folder-files --drive-id b!x --item-id 01ABC');
    }
  });

  it('reads a partner-tenant file on the guest token only, and refuses a tenant id that is not a GUID', async () => {
    if (!command) throw new Error('list-document-comments is not registered');
    const bytes = await buildCommentedXlsx();
    const guest: string[] = [];
    const graph = fakeGraphClient({
      get: () => {
        throw new Error('home token used for a partner-tenant file');
      },
      getBinary: () => {
        throw new Error('home token used for a partner-tenant file');
      },
      getGuest: async (path, tenantId) => {
        guest.push(`get ${path} ${tenantId}`);
        return ok({ name: 'plan.xlsx' });
      },
      getBinaryGuest: async (path, tenantId) => {
        guest.push(`bin ${path} ${tenantId}`);
        return binaryOf(bytes);
      },
    });
    const r = await command.execute(graph, { ...ARGS, tenantId: PARTNER });
    expect(r.ok).toBe(true);
    expect(guest).toEqual([`get /drives/b!x/items/01ABC ${PARTNER}`, `bin /drives/b!x/items/01ABC/content ${PARTNER}`]);
    const bad = await command.execute(graph, { ...ARGS, tenantId: 'contoso' });
    expect(bad.ok).toBe(false);
  });

  it('refuses missing ids as a validation error and passes a failed metadata read through', async () => {
    if (!command) throw new Error('list-document-comments is not registered');
    const missing = await command.execute(fakeGraphClient({}), { driveId: 'b!x' });
    expect(missing.ok).toBe(false);
    if (!missing.ok) expect(missing.error.type).toBe('validation_error');
    const gone = await command.execute(fakeGraphClient({ get: async () => ({ ok: false, error: { type: 'api_error', status: 404, message: 'itemNotFound' } }) }), ARGS);
    expect(gone).toEqual({ ok: false, error: { type: 'api_error', status: 404, message: 'itemNotFound' } });
  });

  it('advertises the drive and item ids, the partner tenant, and the content route', () => {
    if (!command) throw new Error('list-document-comments is not registered');
    expect(command.meta.options.map((o) => o.name)).toEqual(['drive-id', 'item-id', 'tenant-id']);
    expect(command.meta.graphPathTemplate).toBe('/drives/{drive-id}/items/{item-id}/content');
    expect(command.meta.category).toBe('drive');
  });
});

describe('list-document-comments at its edges', () => {
  it('names a nameless folder as such, and reads an item whose folder facet is null as a file', async () => {
    if (!command) throw new Error('list-document-comments is not registered');
    const folder = await command.execute(fakeGraphClient({ get: async () => ok({ folder: {} }) }), ARGS);
    if (!folder.ok) expect(folder.error.message).toStartWith("item '' is a folder");
    const bytes = await buildCommentedXlsx();
    const file = await command.execute(fakeGraphClient({ get: async () => ok({ name: 'plan.xlsx', folder: null }), getBinary: async () => binaryOf(bytes) }), ARGS);
    expect(file.ok).toBe(true);
  });

  it('passes a failed download and an unreadable file type through', async () => {
    if (!command) throw new Error('list-document-comments is not registered');
    const refused = { type: 'api_error', status: 403, message: 'accessDenied' } as const;
    const download = await command.execute(fakeGraphClient({ get: async () => ok({ name: 'plan.xlsx' }), getBinary: async () => ({ ok: false, error: refused }) }), ARGS);
    expect(download).toEqual({ ok: false, error: refused });
    const text = await command.execute(fakeGraphClient({ get: async () => ok({ name: 'notes.txt' }), getBinary: async () => binaryOf(new Uint8Array([104, 105])) }), ARGS);
    expect(text.ok).toBe(false);
    if (!text.ok) expect(text.error).toMatchObject({ status: 415, code: 'unsupported_format' });
  });
});
