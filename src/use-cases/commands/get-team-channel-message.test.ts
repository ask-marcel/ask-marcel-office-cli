import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['get-team-channel-message'];
if (!command) throw new Error('get-team-channel-message is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const params = { teamId: 'tm1', channelId: CHANNEL, messageId: '1700000000000' };
const MESSAGE = `/teams/tm1/channels/${CHANNEL}/messages/1700000000000`;

const capture = (): { paths: string[]; graph: ReturnType<typeof fakeGraphClient> } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ id: '1700000000000', body: { contentType: 'html', content: '<p>hello</p>' } });
    },
  });
  return { paths, graph };
};

describe('getting one channel message', () => {
  it('reads the message under its channel and team', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    expect(paths).toEqual([MESSAGE]);
  });

  it('passes `--select` through, the only query option the item honours', async () => {
    const { paths, graph } = capture();
    await command.execute(graph, { ...params, select: 'id,from,body' });
    expect(paths[0]).toBe(`${MESSAGE}?$select=id%2Cfrom%2Cbody`);
  });

  it('refuses a missing message id before calling Graph', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, { teamId: 'tm1', channelId: CHANNEL });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(paths).toEqual([]);
  });

  it('names the message behind the `403 Forbidden: UnknownError` Graph answers for an unknown message id', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 403, message: 'Forbidden: UnknownError' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.status).toBe(403);
    expect(result.error.code).toBe('cli_rewrite_channel_message_not_found');
    expect(result.error.message).toContain('message-id: "1700000000000"');
    expect(result.error.message).toContain('not a missing scope');
  });

  it('calls a `400 BadRequest: UnknownError` a malformed id, naming both the channel and the message id', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 400, message: 'BadRequest: UnknownError' }) });
    const result = await command.execute(graph, { ...params, messageId: 'abc' });
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_bad_id');
    expect(result.error.message).toContain(`channel-id: "${CHANNEL}", message-id: "abc"`);
  });

  it('advertises team-id, channel-id, message-id and select only: $expand is a 400 on the item', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'channel-id', 'message-id', 'select']);
    expect(command.meta.category).toBe('teams');
    expect(command.meta.pagination).toBeUndefined();
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/channels/{channel-id}/messages/{message-id}');
  });
});
