import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-team-channel-tabs'];
if (!command) throw new Error('list-team-channel-tabs is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const params = { teamId: 'tm1', channelId: CHANNEL };
const TABS = `/teams/tm1/channels/${CHANNEL}/tabs?$expand=teamsApp`;

const capture = (): { paths: string[]; graph: ReturnType<typeof fakeGraphClient> } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ value: [{ id: 'tab1', displayName: 'Files', teamsApp: { id: 'com.microsoft.teamspace.tab.files.sharepoint', displayName: 'Files' } }] });
    },
  });
  return { paths, graph };
};

describe('listing the tabs of a channel', () => {
  it('reads the tabs with the Teams app expanded, so every tab names its app', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    expect(paths).toEqual([TABS]);
  });

  it('appends `--select` and `--filter` after the pinned expand', async () => {
    const { paths, graph } = capture();
    await command.execute(graph, { ...params, select: 'id,displayName', filter: "displayName eq 'Files'" });
    expect(paths[0]).toBe(`${TABS}&$select=id%2CdisplayName&$filter=displayName%20eq%20'Files'`);
  });

  it('names the channel behind the Skype backend failure Graph answers for an unknown channel id', async () => {
    const graph = fakeGraphClient({
      get: async () => err({ type: 'api_error', status: 404, message: 'ItemNotFound: Failed to execute Skype backend request GetThreadRequest.' }),
    });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_not_found');
    expect(result.error.message).toContain(`channel-id: "${CHANNEL}"`);
  });

  it('hands any other Graph error back unchanged', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 404, message: 'ItemNotFound: No Team found with Group id: tm1' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.message).toBe('ItemNotFound: No Team found with Group id: tm1');
  });

  it('advertises team-id, channel-id, select and filter, with the expand pinned in the path and no paging (Graph rejects $top here)', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'channel-id', 'select', 'filter']);
    expect(command.meta.category).toBe('teams');
    expect(command.meta.pagination).toBeUndefined();
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/channels/{channel-id}/tabs?$expand=teamsApp');
  });
});
