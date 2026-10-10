import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { ReadGraph } from '../../infra/read-graph.ts';
import type { ReadCommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ siteId: z.string().min(1) });

const PEOPLE = '$select=id,displayName,mail,userPrincipalName,userType&$top=999';

type Grant = {
  readonly roles?: ReadonlyArray<string>;
  readonly grantedToV2?: { readonly siteGroup?: { readonly displayName?: string } };
  readonly link?: { readonly scope?: string };
};
type Listing = { readonly value?: ReadonlyArray<unknown>; readonly '@odata.nextLink'?: string };

const NO_GROUP =
  "This site belongs to no Microsoft 365 group: its access lives in its SharePoint groups (listed with their roles), and Graph does not list who is in them; the site's owners see them in its settings.";
const LOWER_BOUND =
  "These are the owners and members of the Microsoft 365 group that owns the site. Its SharePoint groups (listed with their roles) can hold more people, or everyone in the organisation, and Graph does not list who is in them: read the group as the fewest people who can open the site; the site's owners see the full list in its settings.";
const TRUNCATED = ' The group has more than 999 members; the first 999 are listed.';

// The library root's permissions name who holds the site beyond its group: the
// SharePoint groups (Owners, Members, Visitors) and any sharing link.
const accessOf = (grants: ReadonlyArray<Grant>): Record<string, unknown> => ({
  sharePointGroups: grants.flatMap((g) => {
    const name = g.grantedToV2?.siteGroup?.displayName;
    return name === undefined ? [] : [{ name, roles: g.roles ?? [] }];
  }),
  sharingLinks: grants.flatMap((g) => (g.link?.scope === undefined ? [] : [{ scope: g.link.scope, roles: g.roles ?? [] }])),
});

const execute = async (graph: ReadGraph, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const drive = await graph.get(`/sites/${parsed.data.siteId}/drive?$select=id,owner`);
  if (!drive.ok) return drive;
  const { id: driveId, owner } = drive.value as {
    readonly id?: string;
    readonly owner?: { readonly group?: { readonly id?: string; readonly displayName?: string; readonly email?: string } };
  };
  const permissions = await graph.get(`/drives/${String(driveId)}/root/permissions`);
  if (!permissions.ok) return permissions;
  const access = accessOf(((permissions.value as Listing).value ?? []) as ReadonlyArray<Grant>);
  const group = owner?.group;
  if (group?.id === undefined) return ok({ group: null, owners: [], members: [], ...access, note: NO_GROUP });
  const [owners, members] = await Promise.all([graph.get(`/groups/${group.id}/owners?${PEOPLE}`), graph.get(`/groups/${group.id}/members?${PEOPLE}`)]);
  if (!owners.ok) return owners;
  if (!members.ok) return members;
  const more = (members.value as Listing)['@odata.nextLink'] === undefined ? '' : TRUNCATED;
  return ok({
    group: { id: group.id, displayName: group.displayName, mail: group.email },
    owners: (owners.value as Listing).value ?? [],
    members: (members.value as Listing).value ?? [],
    ...access,
    note: `${LOWER_BOUND}${more}`,
  });
};

const meta: ReadCommandMeta = {
  summary:
    "Who can open a SharePoint site: the owners and members of the Microsoft 365 group that owns it (name, mail, user type), and the SharePoint groups and sharing links holding its document library, with their roles (`owner`, `write`, `read`). Graph does not list who is inside a SharePoint group, so the group's people are the fewest who can open the site, and the `note` says so. For a page or file that inherits the site's permissions (a wiki page holding credentials, a sensitive document), this is who can read it; an item with its own sharing is not covered. Graph shows a caller who is not a site owner only the grants that apply to them. Find the site id with `search-sharepoint-sites-by-name` or `get-sharepoint-site-by-path`.",
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: "/sites/{site-id}/drive, then its root permissions and the owning group's owners and members",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description:
        'SharePoint site ID (`contoso.sharepoint.com,<guid>,<guid>`), from `search-sharepoint-sites-by-name`, `get-sharepoint-site-by-path` or `search-all-accessible-sites`.',
    },
  ],
  example: "ask-marcel-office list-sharepoint-site-members --site-id 'contoso.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000'",
  responseShape:
    '`{ group: { id, displayName, mail } | null, owners: [user], members: [user], sharePointGroups: [{ name, roles }], sharingLinks: [{ scope, roles }], note }` — each user carries `id`, `displayName`, `mail`, `userPrincipalName`, `userType` (`Member` or `Guest`).',
  effect: 'read',
  scopesRequired: ['Sites.Read.All', 'GroupMember.Read.All'],
};

export { execute, meta, schema };
