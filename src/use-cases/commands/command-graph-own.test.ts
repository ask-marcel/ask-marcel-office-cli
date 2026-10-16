import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeCommand } from '../../test-helpers/command-package-fake.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { withOwnGraph } from './command-graph.ts';
import type { ReadCommand, WriteCommand } from './command-types.ts';

/*
 * Each command runs on the graph its effect allows, whatever graph its caller
 * passes (core, package split phase 3): the full client goes in, and a read
 * command sees the read members only, a write command the write members only.
 */

const membersSeen = (seen: string[][]): ((graph: object) => Promise<ReturnType<typeof ok>>) => {
  const record = async (graph: object): Promise<ReturnType<typeof ok>> => {
    seen.push(Object.keys(graph).toSorted((a, b) => a.localeCompare(b)));
    return ok('ran');
  };
  return record;
};

describe('a command run on the full client', () => {
  it('gets the read view when it reads', async () => {
    const seen: string[][] = [];
    const read: ReadCommand = { ...(fakeCommand() as ReadCommand), execute: membersSeen(seen) };
    expect(await withOwnGraph(read).execute(fakeGraphClient(), {})).toEqual(ok('ran'));
    expect(seen).toEqual([
      ['discoverTenantId', 'fetchUrl', 'get', 'getBinary', 'getBinaryElevated', 'getBinaryGuest', 'getElevated', 'getGuest', 'post', 'teamsChat', 'teamsChatIc3', 'teamsChatMedia'],
    ]);
  });

  it('gets the write view when it writes, and its params', async () => {
    const seen: string[][] = [];
    const params: Array<Record<string, string>> = [];
    const write: WriteCommand = {
      ...(fakeCommand({ effect: 'draft' }) as WriteCommand),
      execute: async (graph, given) => {
        params.push(given);
        return membersSeen(seen)(graph);
      },
    };
    expect(await withOwnGraph(write).execute(fakeGraphClient(), { messageId: 'm1' })).toEqual(ok('ran'));
    expect(seen).toEqual([['delete', 'fetchUrl', 'get', 'getBinary', 'patch', 'post', 'put']]);
    expect(params).toEqual([{ messageId: 'm1' }]);
  });
});
