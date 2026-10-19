import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeCommand } from '../../test-helpers/command-package-fake.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { readCommandOf, writeCommandOf } from './command-graph.ts';
import type { ReadCommand, WriteCommand } from './command-types.ts';

/*
 * A package that holds one kind of command runs each command on the graph of
 * that kind (package split, phase 5): the read package on the read graph, the
 * write package on the write graph. A library caller can still pass a graph
 * with more members, such as the full client; the command gets its view only.
 */

const membersOf = (graph: object): string[] => Object.keys(graph).toSorted((a, b) => a.localeCompare(b));

const READ_MEMBERS = [
  'discoverTenantId',
  'fetchUrl',
  'get',
  'getBinary',
  'getBinaryElevated',
  'getBinaryGuest',
  'getElevated',
  'getGuest',
  'post',
  'teamsChat',
  'teamsChatIc3',
  'teamsChatMedia',
];
const WRITE_MEMBERS = ['delete', 'fetchUrl', 'get', 'getBinary', 'patch', 'post', 'put'];

describe('a command of a package that holds only reads', () => {
  it('runs on the read view of the graph it is given, with its params', async () => {
    const seen: Array<{ readonly members: string[]; readonly params: Record<string, string> }> = [];
    const read: ReadCommand = {
      ...(fakeCommand() as ReadCommand),
      execute: async (graph, params) => {
        seen.push({ members: membersOf(graph), params });
        return ok('listed');
      },
    };
    expect(await readCommandOf(read).execute(fakeGraphClient(), { top: '3' })).toEqual(ok('listed'));
    expect(seen).toEqual([{ members: READ_MEMBERS, params: { top: '3' } }]);
  });

  it("serves a library caller's graph whose members are methods that read `this`", async () => {
    const graph = {
      ...fakeGraphClient(),
      calls: [] as string[],
      async get(path: string) {
        this.calls.push(`get ${path}`);
        return ok({});
      },
      async post(path: string) {
        this.calls.push(`post ${path}`);
        return ok({});
      },
    };
    const read: ReadCommand = {
      ...(fakeCommand() as ReadCommand),
      execute: async (view) => {
        await view.get('/me/drives');
        return view.post('/search/query', {});
      },
    };
    expect((await readCommandOf(read).execute(graph, {})).ok).toBe(true);
    expect(graph.calls).toEqual(['get /me/drives', 'post /search/query']);
  });
});

describe('a command of a package that holds only writes', () => {
  it('runs on the write view of the graph it is given, with its params', async () => {
    const seen: Array<{ readonly members: string[]; readonly params: Record<string, string> }> = [];
    const write: WriteCommand = {
      ...(fakeCommand({ effect: 'draft' }) as WriteCommand),
      execute: async (graph, params) => {
        seen.push({ members: membersOf(graph), params });
        return ok('drafted');
      },
    };
    expect(await writeCommandOf(write).execute(fakeGraphClient(), { messageId: 'm1' })).toEqual(ok('drafted'));
    expect(seen).toEqual([{ members: WRITE_MEMBERS, params: { messageId: 'm1' } }]);
  });

  it("serves a library caller's graph whose members are methods that read `this`", async () => {
    const graph = {
      ...fakeGraphClient(),
      calls: [] as string[],
      async post(path: string) {
        this.calls.push(`post ${path}`);
        return ok({ id: 'd1' });
      },
      async delete(path: string) {
        this.calls.push(`delete ${path}`);
        return ok(1);
      },
    };
    const write: WriteCommand = {
      ...(fakeCommand({ effect: 'draft' }) as WriteCommand),
      execute: async (view) => {
        await view.post('/me/messages', {});
        return view.delete('/me/messages/d1');
      },
    };
    expect(await writeCommandOf(write).execute(graph, {})).toEqual(ok(1));
    expect(graph.calls).toEqual(['post /me/messages', 'delete /me/messages/d1']);
  });
});
