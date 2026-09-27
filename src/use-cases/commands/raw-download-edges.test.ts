import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const run = async (name: string, graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  return command.execute(graph, params);
};

const TEXT = { contentType: 'text/plain', size: 5, base64: Buffer.from('hello').toString('base64') };
const PARTNER = '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04';

describe('download-drive-item-content at its edges', () => {
  it('refuses a folder, named or not, and points at the command that lists what it holds', async () => {
    const named = await run('download-drive-item-content', fakeGraphClient({ get: async () => ok({ name: 'Reports', folder: { childCount: 3 } }) }), {
      driveId: 'd1',
      itemId: 'i1',
    });
    expect(named.ok).toBe(false);
    if (!named.ok) {
      expect(named.error).toMatchObject({ type: 'api_error', status: 400 });
      expect(named.error.message).toContain("item 'Reports' is a folder, not a file — use `list-folder-files --drive-id d1 --item-id i1`");
    }
    const nameless = await run('download-drive-item-content', fakeGraphClient({ get: async () => ok({ folder: {} }) }), { driveId: 'd1', itemId: 'i1' });
    if (!nameless.ok) expect(nameless.error.message).toContain("item '' is a folder");
  });

  it('reads a file whose folder facet is null, and passes a failed metadata read through', async () => {
    const file = await run('download-drive-item-content', fakeGraphClient({ get: async () => ok({ name: 'notes.txt', folder: null }), getBinary: async () => ok(TEXT) }), {
      driveId: 'd1',
      itemId: 'i1',
    });
    expect(file).toEqual({ ok: true, value: { contentType: 'text/plain', size: 5, text: 'hello' } });
    const gone = await run(
      'download-drive-item-content',
      fakeGraphClient({ get: async () => err({ type: 'api_error', status: 404, message: 'itemNotFound: gone', code: 'itemNotFound' }) }),
      {
        driveId: 'd1',
        itemId: 'i1',
      }
    );
    expect(gone.ok).toBe(false);
    if (!gone.ok) expect(gone.error).toMatchObject({ code: 'itemNotFound' });
  });

  it('reads a partner-tenant file on that tenant, and refuses a tenant id that is not a GUID before any call', async () => {
    const seen: string[] = [];
    const guest = fakeGraphClient({
      get: async () => {
        seen.push('home');
        return ok({ name: 'x' });
      },
      getGuest: async (path: string) => {
        seen.push(`guest ${path}`);
        return ok({ name: 'notes.txt' });
      },
      getBinaryGuest: async () => ok(TEXT),
    });
    const read = await run('download-drive-item-content', guest, { driveId: 'd1', itemId: 'i1', tenantId: PARTNER });
    expect(read.ok).toBe(true);
    expect(seen).toEqual(['guest /drives/d1/items/i1']);
    const refused = await run('download-drive-item-content', guest, { driveId: 'd1', itemId: 'i1', tenantId: 'contoso' });
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.error.type).toBe('validation_error');
    expect(seen).toHaveLength(1);
  });
});

describe('list-mail-attachments at its edges', () => {
  it('lists slim metadata by default and keeps a caller --select as given', async () => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      get: async (path: string) => {
        paths.push(decodeURIComponent(path));
        return ok({ value: [] });
      },
    });
    await run('list-mail-attachments', graph, { messageId: 'm1' });
    await run('list-mail-attachments', graph, { messageId: 'm1', select: 'id,name,contentBytes' });
    expect(paths).toEqual(['/me/messages/m1/attachments?$select=id,name,contentType,size,isInline', '/me/messages/m1/attachments?$select=id,name,contentBytes']);
  });

  it('refuses a call without a message id before any Graph call', async () => {
    const refused = await run('list-mail-attachments', fakeGraphClient(), {});
    expect(refused.ok).toBe(false);
    if (!refused.ok) {
      expect(refused.error.type).toBe('validation_error');
      expect(refused.error.message).toContain('--message-id is missing');
    }
  });
});
