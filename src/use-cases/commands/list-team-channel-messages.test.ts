import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-team-channel-messages'];
if (!command) throw new Error('list-team-channel-messages is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const params = { teamId: 'tm1', channelId: CHANNEL };
const MESSAGES = `/teams/tm1/channels/${CHANNEL}/messages`;

const capture = (): { paths: string[]; graph: ReturnType<typeof fakeGraphClient> } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ value: [{ id: '1700000000000', replyToId: null, body: { contentType: 'html', content: '<p>hello</p>' } }] });
    },
  });
  return { paths, graph };
};

describe('listing the messages of a channel', () => {
  it('reads the channel messages collection under the team', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    expect(paths).toEqual([MESSAGES]);
  });

  it('passes `--top`, `--select` and `--expand replies` through', async () => {
    const { paths, graph } = capture();
    await command.execute(graph, { ...params, top: '50', select: 'id,from,body', expand: 'replies' });
    expect(paths[0]).toBe(`${MESSAGES}?$top=50&$select=id%2Cfrom%2Cbody&$expand=replies`);
  });

  it('refuses a page above 50 before calling Graph', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, { ...params, top: '51' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(result.error.message).toContain('at most 50');
    expect(paths).toEqual([]);
  });

  it('names the channel behind the `410 Gone: UnknownError` Graph answers for an unknown channel id', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 410, message: 'Gone: UnknownError' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_not_found');
    expect(result.error.message).toContain(`channel-id: "${CHANNEL}"`);
  });

  it('calls a `400 BadRequest: UnknownError` a malformed channel id, without a message id in play', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 400, message: 'BadRequest: UnknownError' }) });
    const result = await command.execute(graph, { teamId: 'tm1', channelId: 'not-a-channel' });
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_bad_id');
    expect(result.error.message).toContain('channel-id: "not-a-channel"');
    expect(result.error.message).not.toContain('message-id');
  });

  it('hands any other Graph error back unchanged', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 429, message: 'TooManyRequests: slow down' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.message).toBe('TooManyRequests: slow down');
  });

  it('advertises team-id, channel-id, top, select and expand, and pages by nextLink', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'channel-id', 'top', 'select', 'expand']);
    expect(command.meta.options.find((o) => o.name === 'top')?.description).toContain('50');
    expect(command.meta.options.find((o) => o.name === 'expand')?.description).toContain('replies');
    expect(command.meta.category).toBe('teams');
    expect(command.meta.pagination).toBe(true);
    expect(command.meta.paginationStrategy).toBe('nextLinkNoSkip');
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/channels/{channel-id}/messages');
  });
});
