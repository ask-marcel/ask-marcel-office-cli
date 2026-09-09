import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-team-channel-members'];
if (!command) throw new Error('list-team-channel-members is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const params = { teamId: 'tm1', channelId: CHANNEL };

const capture = (): { paths: string[]; graph: ReturnType<typeof fakeGraphClient> } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ value: [{ id: 'm1' }] });
    },
  });
  return { paths, graph };
};

describe('listing the members of a channel', () => {
  it('reads the channel members collection under the team', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    expect(paths).toEqual([`/teams/tm1/channels/${CHANNEL}/members`]);
  });

  it('passes `--select` through', async () => {
    const { paths, graph } = capture();
    await command.execute(graph, { ...params, select: 'id,displayName' });
    expect(paths[0]).toBe(`/teams/tm1/channels/${CHANNEL}/members?$select=id%2CdisplayName`);
  });

  it('names the channel behind the bare `1: NotFound` Graph answers for an unknown channel id', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 404, message: '1: NotFound' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_not_found');
    expect(result.error.message).toContain(`channel-id: "${CHANNEL}"`);
    expect(result.error.message).toContain('list-team-channels');
  });

  it('hands any other Graph error back unchanged', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 400, message: 'BadRequest: teamId needs to be a valid GUID.' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.message).toBe('BadRequest: teamId needs to be a valid GUID.');
  });

  it('advertises team-id, channel-id, filter and select, and pages by nextLink', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'channel-id', 'select', 'filter']);
    expect(command.meta.category).toBe('teams');
    expect(command.meta.pagination).toBe(true);
    expect(command.meta.paginationStrategy).toBe('nextLinkNoSkip');
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/channels/{channel-id}/members');
  });
});
