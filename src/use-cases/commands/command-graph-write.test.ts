// The write case of the registry-graph suite: no registered write command reaches a member
// beyond the basic tier. Moved out of command-graph.test.ts with its text unchanged (package
// split, step 16), so the write package takes this file whole; the helpers are copies of
// that file's.
import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { commands } from './index.ts';

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
const BEYOND_BASIC_MEMBERS = ['getElevated', 'getGuest', 'getBinaryGuest', 'getBinaryElevated', 'discoverTenantId', 'teamsChat', 'teamsChatIc3', 'teamsChatMedia'];

// A full client that records each member called, and the path of each POST.
const recordingClient = (called: string[]): GraphClient => {
  const record = (member: string): (() => Promise<ReturnType<typeof ok<object>>>) => {
    const answer = async (): Promise<ReturnType<typeof ok<object>>> => {
      called.push(member);
      return ok({});
    };
    return answer;
  };
  const members = Object.fromEntries([...new Set([...READ_MEMBERS, ...WRITE_MEMBERS])].map((member) => [member, record(member)]));
  return fakeGraphClient({
    ...members,
    post: async (path: string) => {
      called.push(`post ${path}`);
      return ok({});
    },
  });
};

// A placeholder for each required flag, so a command gets past its own checks.
const placeholderParams = (meta: CommandMeta): Record<string, string> =>
  Object.fromEntries(
    meta.options
      .filter((option) => option.required)
      .map((option) => {
        if (option.argumentHint?.kind === 'iso8601') return [option.key, '2026-01-05T09:00:00Z'];
        if (option.argumentHint?.kind === 'magicValue') return [option.key, option.argumentHint.values[0] ?? 'x'];
        return [option.key, 'x'];
      })
  );

// Runs each command on placeholders against the recording client; answers how
// many of them reached Graph at all, and the names of those whose POST the read
// view refused. The view refuses a write POST before the recording client sees
// it, so only the refusals can show a read command that tried one.
const runEach = async (registered: ReadonlyArray<readonly [string, Command]>, called: string[]): Promise<{ readonly reached: number; readonly refused: ReadonlyArray<string> }> => {
  let reached = 0;
  const refused: string[] = [];
  for (const [name, command] of registered) {
    const before = called.length;
    const result = await command.execute(recordingClient(called), placeholderParams(command.meta));
    if (called.length > before) reached += 1;
    if (!result.ok && result.error.code === 'write_refused') refused.push(name);
  }
  return { reached, refused };
};

describe('the graph the registry gives a command', () => {
  it('never lets a registered write command reach the elevated, guest or chat members', async () => {
    const called: string[] = [];
    const writes = Object.entries(commands).filter(([, command]) => command.meta.effect !== 'read');
    const { reached } = await runEach(writes, called);
    expect(called.filter((member) => BEYOND_BASIC_MEMBERS.includes(member))).toEqual([]);
    expect(reached).toBeGreaterThan(writes.length / 2);
  });
});
