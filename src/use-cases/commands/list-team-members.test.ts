import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-team-members'];
if (!command) throw new Error('list-team-members is not registered');

const capture = (): { paths: string[]; graph: ReturnType<typeof fakeGraphClient> } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ value: [{ id: 'm1', roles: ['owner'] }] });
    },
  });
  return { paths, graph };
};

describe('listing the members of a team', () => {
  it('reads the team members collection', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, { teamId: 'tm1' });
    expect(result.ok).toBe(true);
    expect(paths).toEqual(['/teams/tm1/members']);
  });

  it('passes `--filter` and `--select` through, the two query options Graph honours here', async () => {
    const { paths, graph } = capture();
    await command.execute(graph, { teamId: 'tm1', filter: "(microsoft.graph.aadUserConversationMember/roles/any(r:r eq 'owner'))", select: 'id,displayName,roles' });
    expect(paths[0]).toBe("/teams/tm1/members?$select=id%2CdisplayName%2Croles&$filter=(microsoft.graph.aadUserConversationMember%2Froles%2Fany(r%3Ar%20eq%20'owner'))");
  });

  it('refuses a missing team id before calling Graph', async () => {
    const { paths, graph } = capture();
    const result = await command.execute(graph, {});
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(paths).toEqual([]);
  });

  it('hands a Graph error back unchanged: the team routes already name a bad or unknown team id', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 404, message: 'ItemNotFound: No Team found with Group id: tm1' }) });
    const result = await command.execute(graph, { teamId: 'tm1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.message).toBe('ItemNotFound: No Team found with Group id: tm1');
  });

  it('advertises only team-id, filter and select: Graph rejects $skip and mis-pages on $top here', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'select', 'filter']);
    expect(command.meta.category).toBe('teams');
    expect(command.meta.pagination).toBe(true);
    expect(command.meta.paginationStrategy).toBe('nextLinkNoSkip');
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/members');
  });
});
