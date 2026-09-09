import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-team-channel-message-replies'];
if (!command) throw new Error('list-team-channel-message-replies is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const params = { teamId: 'tm1', channelId: CHANNEL, messageId: '1700000000000' };
const REPLIES = `/teams/tm1/channels/${CHANNEL}/messages/1700000000000/replies`;

const capture = (): { paths: string[]; graph: ReturnType<typeof fakeGraphClient> } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ value: [{ id: '1700000000001', replyToId: '1700000000000' }] });
    },
  });
  return { paths, graph };
};

describe('listing the replies under a channel post', () => {
  it('reads the replies collection under the root post', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    expect(paths).toEqual([REPLIES]);
  });

  it('passes `--top` and `--select` through', async () => {
    const { paths, graph } = capture();
    await command.execute(graph, { ...params, top: '10', select: 'id,from,body' });
    expect(paths[0]).toBe(`${REPLIES}?$top=10&$select=id%2Cfrom%2Cbody`);
  });

  it('refuses a page above 50 before calling Graph', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, { ...params, top: '100' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(paths).toEqual([]);
  });

  it('names both ids behind a `400 BadRequest: UnknownError`', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 400, message: 'BadRequest: UnknownError' }) });
    const result = await command.execute(graph, { ...params, messageId: 'abc' });
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_bad_id');
    expect(result.error.message).toContain('message-id: "abc"');
  });

  it('hands any other Graph error back unchanged', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 404, message: 'ItemNotFound: No Team found with Group id: tm1' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.message).toBe('ItemNotFound: No Team found with Group id: tm1');
  });

  it('advertises team-id, channel-id, message-id, top and select, and pages by nextLink', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'channel-id', 'message-id', 'top', 'select']);
    expect(command.meta.category).toBe('teams');
    expect(command.meta.pagination).toBe(true);
    expect(command.meta.paginationStrategy).toBe('nextLinkNoSkip');
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies');
  });
});
