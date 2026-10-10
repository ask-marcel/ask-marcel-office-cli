import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import type { ReadGraph, ReadOnlyPostPath } from '../../infra/read-graph.ts';
import type { WriteGraph } from '../../infra/write-graph.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import type { Command, CommandMeta, ReadCommand, WriteCommand } from './command-types.ts';
import { commands, registerCommands } from './index.ts';

/*
 * The read guarantee inside the single package (package split, D9): the
 * registry gives each command only the graph its effect allows. A read command
 * holds the read graph (no put, patch, delete or upload session, and a POST
 * that reaches only the two query endpoints); a write command holds the write
 * graph (basic tier only). The type of each command's `execute` says the same
 * at compile time; these tests pin it at run time, through the registry every
 * surface uses.
 */

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
const WRITE_ONLY_MEMBERS = ['patch', 'put', 'delete'];
const BEYOND_BASIC_MEMBERS = ['getElevated', 'getGuest', 'getBinaryGuest', 'getBinaryElevated', 'discoverTenantId', 'teamsChat', 'teamsChatIc3', 'teamsChatMedia'];

const META: CommandMeta = {
  summary: 'probe',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/probe',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/probe',
  options: [],
  example: 'ask-marcel-office probe',
  effect: 'read',
};

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
  it('gives a read command the read graph only: its members, and a POST that refuses any path but the two queries', async () => {
    const posted: string[] = [];
    let seen: ReadGraph | undefined;
    const probe: ReadCommand = {
      schema: z.object({}),
      meta: { ...META, effect: 'read' },
      execute: async (graph) => {
        seen = graph;
        return graph.post('/me/sendMail' as ReadOnlyPostPath, {});
      },
    };
    const full = fakeGraphClient({
      post: async (path: string) => {
        posted.push(path);
        return ok({});
      },
    });
    const result = await registerCommands({ probe }).probe?.execute(full, {});
    expect(Object.keys(seen ?? {}).toSorted((a, b) => a.localeCompare(b))).toEqual(READ_MEMBERS);
    expect(result?.ok).toBe(false);
    expect(posted).toEqual([]);
  });

  it('gives a write command the write graph only: the basic tier and the four writes', async () => {
    let seen: WriteGraph | undefined;
    const probe: WriteCommand = {
      schema: z.object({}),
      meta: { ...META, graphMethod: 'POST', effect: 'draft' },
      execute: async (graph) => {
        seen = graph;
        return ok({});
      },
    };
    await registerCommands({ probe }).probe?.execute(fakeGraphClient(), {});
    expect(Object.keys(seen ?? {}).toSorted((a, b) => a.localeCompare(b))).toEqual(WRITE_MEMBERS);
  });

  it("serves a library caller's graph whose members are methods that read `this`, for a read command and a write command alike", async () => {
    const graph = {
      ...fakeGraphClient(),
      calls: [] as string[],
      async get(path: string) {
        this.calls.push(`get ${path}`);
        return ok({ value: [] });
      },
      async post(path: string) {
        this.calls.push(`post ${path}`);
        return ok({ id: 'draft-1' });
      },
    };
    const listed = await commands['list-drives']?.execute(graph, {});
    const drafted = await commands['create-mail-draft']?.execute(graph, { subject: 'Hi', bodyContent: 'Hello', toRecipients: 'a@example.com' });
    expect([listed?.ok, drafted?.ok]).toEqual([true, true]);
    expect(graph.calls).toEqual(['get /me/drives', 'post /me/messages']);
  });

  it('never lets a registered read command call a member that writes, or POST anywhere but the two queries', async () => {
    const called: string[] = [];
    const reads = Object.entries(commands).filter(([, command]) => command.meta.effect === 'read');
    const { reached, refused } = await runEach(reads, called);
    expect(called.filter((member) => WRITE_ONLY_MEMBERS.includes(member))).toEqual([]);
    expect(refused).toEqual([]);
    // Most read commands get past their own checks on placeholders and reach
    // Graph, so the checks above are not empty ones.
    expect(reached).toBeGreaterThan(reads.length / 2);
  });

  it('never lets a registered write command reach the elevated, guest or chat members', async () => {
    const called: string[] = [];
    const writes = Object.entries(commands).filter(([, command]) => command.meta.effect !== 'read');
    const { reached } = await runEach(writes, called);
    expect(called.filter((member) => BEYOND_BASIC_MEMBERS.includes(member))).toEqual([]);
    expect(reached).toBeGreaterThan(writes.length / 2);
  });
});
