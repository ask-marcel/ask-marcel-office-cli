import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-team-channel-messages-delta'];
if (!command) throw new Error('list-team-channel-messages-delta is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const DELTA = `/teams/tm1/channels/${CHANNEL}/messages/delta`;
const params = { teamId: 'tm1', channelId: CHANNEL };

const capture = (): { paths: string[]; graph: ReturnType<typeof fakeGraphClient> } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ value: [], '@odata.deltaLink': `https://graph.microsoft.com/v1.0${DELTA}?$deltatoken=t` });
    },
  });
  return { paths, graph };
};

describe('tracking what changed in a channel', () => {
  it('starts a full sync when no --since is given', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    expect(paths).toEqual([DELTA]);
  });

  it('bounds the first sync with --since as a lastModifiedDateTime filter, with --top and --expand replies after it', async () => {
    const { paths, graph } = capture();
    await command.execute(graph, { ...params, since: '2026-09-01T00:00:00Z', top: '5', expand: 'replies' });
    expect(paths[0]).toBe(`${DELTA}?$filter=lastModifiedDateTime%20gt%202026-09-01T00:00:00Z&$top=5&$expand=replies`);
  });

  it('accepts the relative date vocabulary for --since', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, { ...params, since: '7d' });
    expect(result.ok).toBe(true);
    expect(paths[0]?.startsWith(`${DELTA}?$filter=lastModifiedDateTime%20gt%2020`)).toBe(true);
  });

  it('refuses a page above 50 and an unreadable date before calling Graph', async () => {
    const rejected: ReadonlyArray<Record<string, string>> = [{ top: '51' }, { since: 'whenever' }];
    for (const extra of rejected) {
      const { paths, graph } = capture();
      const result = await command.execute(graph, { ...params, ...extra });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('validation_error');
      expect(paths).toEqual([]);
    }
  });

  it('names the channel behind the `410 Gone: UnknownError` an unknown channel answers', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 410, message: 'Gone: UnknownError' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_not_found');
  });

  it('advertises team-id, channel-id, since, top and expand, and pages by nextLink then deltaLink', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'channel-id', 'since', 'top', 'expand']);
    expect(command.meta.category).toBe('teams');
    expect(command.meta.pagination).toBe(true);
    expect(command.meta.paginationStrategy).toBe('deltaLink');
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/channels/{channel-id}/messages/delta?$filter=lastModifiedDateTime gt {since}');
  });
});
