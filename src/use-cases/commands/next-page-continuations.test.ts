import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const GRAPH = 'https://graph.microsoft.com/v1.0';

type Seen = { path: string; headers: Record<string, string> | undefined };

const recording = (page: unknown, seen: Seen[]): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    get: async (path: string, headers?: Record<string, string>): Promise<Result<unknown, GraphError>> => {
      seen.push({ path, headers });
      return ok(page);
    },
  });

const nextPage = async (graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands['next-page'];
  if (!command) throw new Error('next-page is not registered');
  return command.execute(graph, params);
};

describe('next-page on a To Do task listing', () => {
  it('gives every task of a later page the link the To Do web app opens it with, as page one does', async () => {
    const seen: Seen[] = [];
    const page = { value: [{ id: 'AAMkT1=', title: 'Renew badge' }], '@odata.nextLink': `${GRAPH}/me/todo/lists/L1/tasks?$skip=20` };
    const result = await nextPage(recording(page, seen), { url: `${GRAPH}/me/todo/lists/L1/tasks?$skip=10` });
    if (!result.ok) throw new Error(result.error.message);
    expect((result.value as { value: ReadonlyArray<{ webUrl?: string }> }).value[0]?.webUrl).toBe('https://to-do.office.com/tasks/id/AAMkT1=/details');
  });

  it('links the live tasks of a delta page but not the deleted ones, which no longer open', async () => {
    const page = {
      value: [
        { id: 'T2', title: 'Book room' },
        { id: 'T3', '@removed': { reason: 'deleted' } },
      ],
    };
    const result = await nextPage(recording(page, []), { url: `${GRAPH}/me/todo/lists/L1/tasks/delta?$deltatoken=abc` });
    if (!result.ok) throw new Error(result.error.message);
    const [live, removed] = (result.value as { value: ReadonlyArray<Record<string, unknown>> }).value;
    expect(live?.['webUrl']).toBe('https://to-do.office.com/tasks/id/T2/details');
    expect(removed).not.toHaveProperty('webUrl');
  });

  it('links a task page read from a cursor that carries no query at all', async () => {
    const result = await nextPage(recording({ value: [{ id: 'T9' }] }, []), { url: `${GRAPH}/me/todo/lists/L1/tasks` });
    if (!result.ok) throw new Error(result.error.message);
    expect((result.value as { value: ReadonlyArray<{ webUrl?: string }> }).value[0]?.webUrl).toBe('https://to-do.office.com/tasks/id/T9/details');
  });

  it('leaves a page of anything else as Graph sent it', async () => {
    const page = { value: [{ id: 'M1', subject: 'Budget' }] };
    const result = await nextPage(recording(page, []), { url: `${GRAPH}/me/messages?$skiptoken=x` });
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toEqual(page);
  });
});

describe('next-page --top', () => {
  it('asks for that page size on the continuation, since Graph honours a page-size preference only on the request that carries it', async () => {
    const seen: Seen[] = [];
    await nextPage(recording({ value: [] }, seen), { url: `${GRAPH}/me/mailFolders/archive/messages/delta?$skiptoken=x`, top: '100' });
    expect(seen).toEqual([{ path: '/me/mailFolders/archive/messages/delta?$skiptoken=x', headers: { Prefer: 'odata.maxpagesize=100' } }]);
    const plain: Seen[] = [];
    await nextPage(recording({ value: [] }, plain), { url: `${GRAPH}/me/mailFolders/archive/messages/delta?$skiptoken=x` });
    expect(plain[0]?.headers).toBeUndefined();
  });

  it('is refused on a chat or partner-tenant cursor, whose page size it cannot carry, and past the 1,000 cap', async () => {
    const refusals: ReadonlyArray<Record<string, string>> = [
      { url: `${GRAPH}/me/chats?$skiptoken=x`, top: '50' },
      { url: `${GRAPH}/drives/b!x/items/i1/children?$skiptoken=x`, top: '50', tenantId: '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04' },
      { url: `${GRAPH}/me/messages?$skiptoken=x`, top: '1001' },
    ];
    for (const params of refusals) {
      const result = await nextPage(recording({ value: [] }, []), params);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('validation_error');
    }
    const chat = await nextPage(recording({ value: [] }, []), { url: `${GRAPH}/me/chats?$skiptoken=x`, top: '50' });
    if (!chat.ok) expect(chat.error.message).toContain('--top applies to a cursor read on the basic token');
  });
});
