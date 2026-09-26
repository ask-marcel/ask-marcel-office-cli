import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const SELECT = '$select=id,displayName,mail,userPrincipalName,userType&$top=999';
const grant = (role: string, name: string): Record<string, unknown> => ({ roles: [role], grantedToV2: { siteGroup: { displayName: name, id: '3', loginName: name } } });
const person = (name: string, mail: string): Record<string, unknown> => ({
  '@odata.type': '#microsoft.graph.user',
  id: mail,
  displayName: name,
  mail,
  userPrincipalName: mail,
  userType: 'Member',
});

const siteGraph = (answers: Record<string, Result<unknown, GraphError>>): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({ get: async (path: string) => answers[path] ?? err({ type: 'api_error', status: 404, message: `itemNotFound: ${path}` }) });

const GROUP_SITE = {
  '/sites/s1/drive?$select=id,owner': ok({ id: 'd1', owner: { group: { id: 'g1', displayName: 'IT wiki', email: 'itwiki@contoso.example' } } }),
  '/drives/d1/root/permissions': ok({
    value: [grant('owner', 'IT wiki Owners'), grant('read', 'IT wiki Visitors'), grant('write', 'IT wiki Members'), { roles: ['read'], link: { scope: 'organization' } }],
  }),
  [`/groups/g1/owners?${SELECT}`]: ok({ value: [person('Robin Chen', 'robin.chen@contoso.example')] }),
  [`/groups/g1/members?${SELECT}`]: ok({ value: [person('Robin Chen', 'robin.chen@contoso.example'), person('Alex Kim', 'alex.kim@contoso.example')] }),
};

const run = async (graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands['list-sharepoint-site-members'];
  if (!command) throw new Error('list-sharepoint-site-members is not registered');
  return command.execute(graph, params);
};

const LOWER_BOUND =
  "These are the owners and members of the Microsoft 365 group that owns the site. Its SharePoint groups (listed with their roles) can hold more people, or everyone in the organisation, and Graph does not list who is in them: read the group as the fewest people who can open the site; the site's owners see the full list in its settings.";

describe('list-sharepoint-site-members', () => {
  it("answers the owning group's owners and members, and the SharePoint groups holding the library with their roles", async () => {
    const result = await run(siteGraph(GROUP_SITE), { siteId: 's1' });
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toEqual({
      group: { id: 'g1', displayName: 'IT wiki', mail: 'itwiki@contoso.example' },
      owners: [person('Robin Chen', 'robin.chen@contoso.example')],
      members: [person('Robin Chen', 'robin.chen@contoso.example'), person('Alex Kim', 'alex.kim@contoso.example')],
      sharePointGroups: [
        { name: 'IT wiki Owners', roles: ['owner'] },
        { name: 'IT wiki Visitors', roles: ['read'] },
        { name: 'IT wiki Members', roles: ['write'] },
      ],
      sharingLinks: [{ scope: 'organization', roles: ['read'] }],
      note: LOWER_BOUND,
    });
  });

  it('says so when a site has no group, and when a group lists more than one page of members', async () => {
    const noGroup = await run(
      siteGraph({
        ...GROUP_SITE,
        '/sites/s1/drive?$select=id,owner': ok({ id: 'd1', owner: { user: { displayName: 'System Account' } } }),
        '/drives/d1/root/permissions': ok({ value: [{ grantedToV2: { siteGroup: { displayName: 'Board Visitors' } } }] }),
      }),
      { siteId: 's1' }
    );
    if (!noGroup.ok) throw new Error(noGroup.error.message);
    expect(noGroup.value).toMatchObject({
      group: null,
      owners: [],
      members: [],
      sharePointGroups: [{ name: 'Board Visitors', roles: [] }],
      sharingLinks: [],
      note: "This site belongs to no Microsoft 365 group: its access lives in its SharePoint groups (listed with their roles), and Graph does not list who is in them; the site's owners see them in its settings.",
    });
    const ownerless = await run(
      siteGraph({
        ...GROUP_SITE,
        '/sites/s1/drive?$select=id,owner': ok({ id: 'd1' }),
        '/drives/d1/root/permissions': ok({ value: [{ roles: ['write'], grantedToV2: { user: { displayName: 'Robin Chen' } } }] }),
      }),
      { siteId: 's1' }
    );
    if (!ownerless.ok) throw new Error(ownerless.error.message);
    expect(ownerless.value).toMatchObject({ group: null, sharePointGroups: [], sharingLinks: [] });
    const big = await run(
      siteGraph({
        ...GROUP_SITE,
        [`/groups/g1/members?${SELECT}`]: ok({ value: [person('Alex Kim', 'alex.kim@contoso.example')], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next' }),
      }),
      { siteId: 's1' }
    );
    if (!big.ok) throw new Error(big.error.message);
    expect((big.value as { note: string }).note).toBe(`${LOWER_BOUND} The group has more than 999 members; the first 999 are listed.`);
  });

  it('passes on a failed read at any step, and refuses a call without a site', async () => {
    for (const path of ['/sites/s1/drive?$select=id,owner', '/drives/d1/root/permissions', `/groups/g1/owners?${SELECT}`, `/groups/g1/members?${SELECT}`]) {
      const answers = { ...GROUP_SITE, [path]: err({ type: 'api_error' as const, status: 403, message: 'accessDenied: no' }) };
      const result = await run(siteGraph(answers), { siteId: 's1' });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.message).toBe('accessDenied: no');
    }
    const refused = await run(siteGraph(GROUP_SITE), {});
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.error.type).toBe('validation_error');
  });
});
