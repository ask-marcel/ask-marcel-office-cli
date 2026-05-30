import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { searchIndexTotal } from './search-index-total.ts';

/**
 * Enumerate every SharePoint / Teams / OneDrive document library the signed-in
 * user can reach. No single Graph endpoint answers this, so the command unions
 * the vectors the delegated Teams token can hit and dedupes by drive id:
 *
 *   - `/me/drives`                      → personal OneDrive(s)        [personal]
 *   - `/me/joinedTeams` → group drive   → Teams document libraries    [joinedTeam]
 *   - `/me/memberOf` (Unified groups) → group drive → SharePoint M365-group sites [memberOfGroup]
 *   - `/me/drive/sharedWithMe`          → drives behind shared items  [sharedWithMe]
 *                                         (catches direct-link-only sites the
 *                                          tenant search index misses)
 *
 * Each drive is tagged with the source(s) that surfaced it. Failures of a
 * single vector or a single group's drive do not fail the command — they land
 * in `partialErrors[]` (a group with no provisioned drive 404s and is dropped
 * silently). `/me/followedSites` and `/sites/delta()` are deliberately avoided
 * (they 403 on this token).
 */

const GRAPH_PREFIX = 'https://graph.microsoft.com/v1.0';
const DEFAULT_MAX_GROUPS = 100;
const MAX_PAGES = 25;

type Source = 'activity' | 'channel' | 'joinedTeam' | 'memberOfGroup' | 'personal' | 'sharedWithMe' | 'siteLibrary';
type DriveResource = { readonly id?: string; readonly name?: string; readonly driveType?: string; readonly webUrl?: string };
type Accumulator = { name: string; driveType: string; webUrl: string; groupId: string | null; readonly sources: Set<Source> };

const schema = z.object({
  maxGroups: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
});

const stripPrefix = (nextLink: string): string => (nextLink.startsWith(`${GRAPH_PREFIX}/`) ? nextLink.slice(GRAPH_PREFIX.length) : nextLink);

// Follow @odata.nextLink (bounded) and concatenate every `value[]` page.
const listAll = async (graph: GraphClient, firstPath: string): Promise<Result<ReadonlyArray<unknown>, GraphError>> => {
  const items: Array<unknown> = [];
  let path: string | undefined = firstPath;
  for (let page = 0; page < MAX_PAGES && path !== undefined; page += 1) {
    const res = await graph.get(path);
    if (!res.ok) return res;
    const body = res.value as { value?: ReadonlyArray<unknown>; '@odata.nextLink'?: string };
    if (Array.isArray(body.value)) items.push(...body.value);
    const next = body['@odata.nextLink'];
    path = typeof next === 'string' ? stripPrefix(next) : undefined;
  }
  return ok(items);
};

const idOf = (o: unknown): string | undefined => {
  if (o === null || typeof o !== 'object') return undefined;
  const id = (o as { id?: unknown }).id;
  return typeof id === 'string' ? id : undefined;
};

const isUnifiedGroup = (g: unknown): boolean => {
  if (g === null || typeof g !== 'object') return false;
  const types = (g as { groupTypes?: unknown }).groupTypes;
  return Array.isArray(types) && types.includes('Unified');
};

const sharedDriveIds = (body: unknown): ReadonlyArray<string> => {
  const items = (body as { value?: ReadonlyArray<unknown> } | null)?.value;
  if (!Array.isArray(items)) return [];
  const ids: Array<string> = [];
  for (const item of items) {
    if (item === null || typeof item !== 'object') continue;
    const driveId = (item as { remoteItem?: { parentReference?: { driveId?: unknown } } }).remoteItem?.parentReference?.driveId;
    if (typeof driveId === 'string') ids.push(driveId);
  }
  return ids;
};

const isNoDrive = (e: GraphError): boolean => e.type === 'api_error' && e.status === 404;

const upsert = (drives: Map<string, Accumulator>, drive: DriveResource, source: Source, groupId: string | null): void => {
  const id = drive.id;
  if (typeof id !== 'string') return;
  const cur = drives.get(id);
  if (cur === undefined) {
    drives.set(id, { name: drive.name ?? '', driveType: drive.driveType ?? '', webUrl: drive.webUrl ?? '', groupId, sources: new Set([source]) });
    return;
  }
  cur.sources.add(source);
  if (cur.groupId === null && groupId !== null) cur.groupId = groupId;
  if (cur.name === '' && typeof drive.name === 'string') cur.name = drive.name;
  if (cur.driveType === '' && typeof drive.driveType === 'string') cur.driveType = drive.driveType;
  if (cur.webUrl === '' && typeof drive.webUrl === 'string') cur.webUrl = drive.webUrl;
};

// Private/shared channels keep their files in their own SharePoint sites, not the
// parent team's default drive — so only those need separate enumeration.
const isPrivateOrShared = (ch: unknown): boolean => {
  if (ch === null || typeof ch !== 'object') return false;
  const t = (ch as { membershipType?: unknown }).membershipType;
  return t === 'private' || t === 'shared';
};

const folderDriveId = (body: unknown): string | undefined => {
  const driveId = (body as { parentReference?: { driveId?: unknown } } | null)?.parentReference?.driveId;
  return typeof driveId === 'string' ? driveId : undefined;
};

type PartialErrors = Array<{ readonly source: string; readonly error: GraphError }>;

// Add `source` to drives already discovered by another vector.
const tagKnown = (drives: Map<string, Accumulator>, ids: ReadonlyArray<string>, source: Source): void => {
  for (const id of new Set(ids)) {
    const cur = drives.get(id);
    if (cur !== undefined) cur.sources.add(source);
  }
};

// Resolve `/drives/{id}` for ids not yet known, capped at `maxGroups`. Returns true if the cap was hit.
const enrichUnknownDrives = async (
  graph: GraphClient,
  ids: ReadonlyArray<string>,
  source: Source,
  drives: Map<string, Accumulator>,
  maxGroups: number,
  partialErrors: PartialErrors
): Promise<boolean> => {
  const unknown = [...new Set(ids)].filter((id) => !drives.has(id));
  const capped = unknown.slice(0, maxGroups);
  const results = await Promise.all(capped.map((id) => graph.get(`/drives/${id}`)));
  capped.forEach((id, i) => {
    const r = results[i];
    if (r === undefined) return;
    if (r.ok) upsert(drives, r.value as DriveResource, source, null);
    else if (!isNoDrive(r.error)) partialErrors.push({ source: `/drives/${id}`, error: r.error });
  });
  return unknown.length > maxGroups;
};

// Collect drive ids behind every private/shared channel of the given teams: list channels,
// then resolve each non-standard channel's files folder to its backing drive id.
const collectChannelDriveIds = async (
  graph: GraphClient,
  teamIds: ReadonlyArray<string>,
  maxGroups: number,
  partialErrors: PartialErrors
): Promise<{ driveIds: ReadonlyArray<string>; truncated: boolean }> => {
  const lists = await Promise.all(teamIds.map((id) => listAll(graph, `/teams/${id}/channels?$select=id,membershipType`)));
  const refs: Array<{ teamId: string; channelId: string }> = [];
  teamIds.forEach((teamId, i) => {
    const r = lists[i];
    if (r === undefined) return;
    if (!r.ok) {
      if (!isNoDrive(r.error)) partialErrors.push({ source: `/teams/${teamId}/channels`, error: r.error });
      return;
    }
    for (const ch of r.value) {
      const channelId = isPrivateOrShared(ch) ? idOf(ch) : undefined;
      if (channelId !== undefined) refs.push({ teamId, channelId });
    }
  });
  const capped = refs.slice(0, maxGroups);
  const folders = await Promise.all(capped.map((c) => graph.get(`/teams/${c.teamId}/channels/${c.channelId}/filesFolder`)));
  const driveIds = new Set<string>();
  capped.forEach((c, i) => {
    const r = folders[i];
    if (r === undefined) return;
    if (!r.ok) {
      if (!isNoDrive(r.error)) partialErrors.push({ source: `/teams/${c.teamId}/channels/${c.channelId}/filesFolder`, error: r.error });
      return;
    }
    const driveId = folderDriveId(r.value);
    if (driveId !== undefined) driveIds.add(driveId);
  });
  return { driveIds: [...driveIds], truncated: refs.length > maxGroups };
};

// Activity signals: drives behind files the user recently used / follows / that trend around
// them. Surfaces secondary libraries and active sites that pure membership/sharing can miss.
const ACTIVITY_PATHS = ['/me/drive/recent', '/me/drive/following', '/me/insights/trending', '/me/insights/used', '/me/insights/shared'] as const;

// driveItems carry parentReference.driveId; insights items carry resourceReference.id = "drives/{driveId}/items/…".
const itemDriveId = (item: unknown): string | undefined => {
  if (item === null || typeof item !== 'object') return undefined;
  const parentId = (item as { parentReference?: { driveId?: unknown } }).parentReference?.driveId;
  if (typeof parentId === 'string') return parentId;
  const refId = (item as { resourceReference?: { id?: unknown } }).resourceReference?.id;
  if (typeof refId !== 'string') return undefined;
  const match = refId.match(/(?:^|\/)drives\/(b![^/]+)/);
  return match === null ? undefined : match[1];
};

const collectActivityDriveIds = async (graph: GraphClient, partialErrors: PartialErrors): Promise<ReadonlyArray<string>> => {
  const results = await Promise.all(ACTIVITY_PATHS.map((p) => listAll(graph, p)));
  const ids = new Set<string>();
  ACTIVITY_PATHS.forEach((p, i) => {
    const r = results[i];
    if (r === undefined) return;
    if (!r.ok) {
      if (!isNoDrive(r.error)) partialErrors.push({ source: p, error: r.error });
      return;
    }
    for (const item of r.value) {
      const id = itemDriveId(item);
      if (id !== undefined) ids.add(id);
    }
  });
  return [...ids];
};

// The group/site vectors only surface a site's DEFAULT library. Path-address each known
// SharePoint site (from a drive's webUrl) and enumerate ALL its document libraries.
const siteAddress = (webUrl: string): string | undefined => {
  const match = webUrl.match(/^https:\/\/([^/]+)\/sites\/([^/]+)/);
  return match === null ? undefined : `${match[1]}:/sites/${match[2]}:`;
};

const driveValues = (body: unknown): ReadonlyArray<unknown> => {
  const value = (body as { value?: ReadonlyArray<unknown> } | null)?.value;
  return Array.isArray(value) ? value : [];
};

const addSiteLibraries = async (graph: GraphClient, drives: Map<string, Accumulator>, maxGroups: number, partialErrors: PartialErrors): Promise<boolean> => {
  const sites = new Set<string>();
  for (const acc of drives.values()) {
    const addr = siteAddress(acc.webUrl);
    if (addr !== undefined) sites.add(addr);
  }
  const capped = [...sites].slice(0, maxGroups);
  const results = await Promise.all(capped.map((addr) => graph.get(`/sites/${addr}/drives`)));
  capped.forEach((addr, i) => {
    const r = results[i];
    if (r === undefined) return;
    if (!r.ok) {
      if (!isNoDrive(r.error)) partialErrors.push({ source: `/sites/${addr}/drives`, error: r.error });
      return;
    }
    for (const lib of driveValues(r.value)) {
      const libId = idOf(lib);
      if (libId !== undefined && !drives.has(libId)) upsert(drives, lib as DriveResource, 'siteLibrary', null);
    }
  });
  return sites.size > maxGroups;
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const maxGroups = Number(parsed.data.maxGroups ?? String(DEFAULT_MAX_GROUPS));

  const [personalR, teamsR, groupsR, sharedR] = await Promise.all([
    listAll(graph, '/me/drives'),
    listAll(graph, '/me/joinedTeams'),
    listAll(graph, '/me/memberOf/microsoft.graph.group?$select=id,displayName,groupTypes'),
    graph.get('/me/drive/sharedWithMe'),
  ]);

  const partialErrors: Array<{ readonly source: string; readonly error: GraphError }> = [];
  const recordRoot = (source: string, r: Result<unknown, GraphError>): void => {
    if (!r.ok) partialErrors.push({ source, error: r.error });
  };
  recordRoot('/me/drives', personalR);
  recordRoot('/me/joinedTeams', teamsR);
  recordRoot('/me/memberOf', groupsR);
  recordRoot('/me/drive/sharedWithMe', sharedR);

  const drives = new Map<string, Accumulator>();
  if (personalR.ok) for (const d of personalR.value) upsert(drives, d as DriveResource, 'personal', null);

  const teamIds = new Set<string>();
  if (teamsR.ok)
    for (const t of teamsR.value) {
      const id = idOf(t);
      if (id !== undefined) teamIds.add(id);
    }
  const memberIds = new Set<string>();
  if (groupsR.ok)
    for (const g of groupsR.value) {
      const id = isUnifiedGroup(g) ? idOf(g) : undefined;
      if (id !== undefined) memberIds.add(id);
    }

  const allGroupIds = [...new Set([...teamIds, ...memberIds])];
  const groupIds = allGroupIds.slice(0, maxGroups);
  let truncated = allGroupIds.length > maxGroups;

  const groupDriveResults = await Promise.all(groupIds.map((id) => graph.get(`/groups/${id}/drive`)));
  groupIds.forEach((id, i) => {
    const r = groupDriveResults[i];
    if (r === undefined) return;
    if (r.ok) {
      if (teamIds.has(id)) upsert(drives, r.value as DriveResource, 'joinedTeam', id);
      if (memberIds.has(id)) upsert(drives, r.value as DriveResource, 'memberOfGroup', id);
    } else if (!isNoDrive(r.error)) {
      partialErrors.push({ source: `/groups/${id}/drive`, error: r.error });
    }
  });

  const sharedIds = sharedR.ok ? sharedDriveIds(sharedR.value) : [];
  tagKnown(drives, sharedIds, 'sharedWithMe');
  if (await enrichUnknownDrives(graph, sharedIds, 'sharedWithMe', drives, maxGroups, partialErrors)) truncated = true;

  // Fifth vector: private/shared channel drives (their own sites, not /groups/{id}/drive).
  const channelTeamIds = groupIds.filter((id) => teamIds.has(id));
  const channels = await collectChannelDriveIds(graph, channelTeamIds, maxGroups, partialErrors);
  tagKnown(drives, channels.driveIds, 'channel');
  const channelCapHit = await enrichUnknownDrives(graph, channels.driveIds, 'channel', drives, maxGroups, partialErrors);
  if (channels.truncated || channelCapHit) truncated = true;

  // Sixth vector: drives behind recently-used / followed / trending items (activity signals).
  const activityIds = await collectActivityDriveIds(graph, partialErrors);
  tagKnown(drives, activityIds, 'activity');
  if (await enrichUnknownDrives(graph, activityIds, 'activity', drives, maxGroups, partialErrors)) truncated = true;

  // Seventh vector: every OTHER document library of each known SharePoint site (not just the default).
  if (await addSiteLibraries(graph, drives, maxGroups, partialErrors)) truncated = true;

  const value = [...drives.entries()]
    .map(([id, d]) => ({ id, name: d.name, driveType: d.driveType, webUrl: d.webUrl, sources: [...d.sources].sort(), ...(d.groupId !== null ? { groupId: d.groupId } : {}) }))
    .sort((a, b) => a.id.localeCompare(b.id));

  if (value.length === 0 && partialErrors.length > 0) return err(partialErrors[0]?.error ?? { type: 'api_error', status: 500, message: 'all drive-discovery sub-requests failed' });

  // Best-effort index-wide file count (driveItems) for context — NOT scoped to the drives above.
  const fileEstimate = await searchIndexTotal(graph, 'driveItem');

  return ok({
    value,
    count: value.length,
    ...(fileEstimate !== undefined ? { fileEstimate } : {}),
    ...(truncated ? { truncated: true } : {}),
    ...(partialErrors.length > 0 ? { partialErrors } : {}),
  });
};

const meta: CommandMeta = {
  summary:
    'Enumerate every drive (document library) the signed-in user can reach — personal OneDrive(s), Teams libraries, SharePoint M365-group sites, drives behind files shared with the user, private/shared Teams channel sites, drives behind recently-used / followed / trending items (activity signals), AND every NON-default document library of each discovered SharePoint site — by unioning `/me/drives`, `/me/joinedTeams`, `/me/memberOf` (Unified groups → `/groups/{id}/drive`), `/me/drive/sharedWithMe`, per-team `/teams/{id}/channels` → `/channels/{ch}/filesFolder` (private/shared channels only — their files live in their own site, not the team default drive), `/me/drive/recent` + `/me/drive/following` + `/me/insights/{trending,used,shared}`, and a path-addressed `/sites/{host}:/sites/{name}:/drives` per discovered site (catches secondary libraries like "Teams Wiki Data" the default-drive vectors miss). Unlike `search-sharepoint-sites-by-name` (which relies on the tenant search index and misses direct-link-only sites + OneDrives), these vectors surface drives the search index never returns; the index in turn returns sites you can open but are not a member of, so the *union of both commands* is the practical maximum on a delegated token. Each drive is tagged with the `sources[]` that found it (a drive can have several). Per-vector / per-group failures do not fail the command — they appear in `partialErrors[]`; a group/channel/site with no provisioned drive is dropped silently. Fans out one `/groups/{id}/drive` + one `/teams/{id}/channels` call per joined team + member group, a `filesFolder` call per private/shared channel, five fixed activity calls, and one `/sites/{id}/drives` call per discovered site (all capped by `--max-groups`, default 100; raise carefully — large memberships can hit 429 throttling). `/me/followedSites` is not used — it 403s on this token.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate:
    '/me/drives + /me/joinedTeams + /me/memberOf + /me/drive/sharedWithMe + per-group /groups/<id>/drive + per-team /teams/<id>/channels/<ch>/filesFolder + /me/drive/recent + /me/drive/following + /me/insights/<trending|used|shared> + per-site /sites/<host>:/sites/<name>:/drives',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-list',
  options: [
    {
      name: 'max-groups',
      key: 'maxGroups',
      required: false,
      description:
        'Safety cap on each fan-out (positive integer; default 100): the per-group `/groups/{id}/drive` calls, the per-team `/teams/{id}/channels` enumeration, the per-private/shared-channel `filesFolder` lookups, the per-site `/sites/{id}/drives` enumeration, and the `/drives/{id}` enrichment for shared-, channel-, and activity-only drives. A user in hundreds of teams/groups/sites would otherwise issue hundreds of parallel requests (429 risk). When any cap is hit the response carries `truncated: true`.',
      argumentHint: { kind: 'magicValue', values: ['100'] },
    },
  ],
  example: 'ask-marcel list-accessible-drives --output json',
  responseShape:
    '`{ value: [{ id, name, driveType, webUrl, sources: ["activity"|"channel"|"joinedTeam"|"memberOfGroup"|"personal"|"sharedWithMe"|"siteLibrary"], groupId? }], count, fileEstimate?, truncated?: true, partialErrors?: [{ source, error }] }`. `value[]` is deduped by drive `id` and sorted by id; `sources[]` lists every vector that surfaced the drive (`channel` = a private/shared Teams channel files folder; `activity` = a recently-used / followed / trending item drive; `siteLibrary` = a non-default document library of a discovered site); `groupId` is present only for Teams/group drives. `fileEstimate` (best-effort, omitted if the extra query fails) is the Microsoft Search index\'s security-trimmed `driveItem` count — roughly how many files+folders you can access across ALL of SharePoint/OneDrive; it is INDEX-WIDE, not limited to the `value[]` drives above. `truncated: true` means a `--max-groups` cap was hit — raise it to see more. `partialErrors[]` (when present) names each vector, group, channel, or site whose sub-call failed.',
};

export { execute, meta, schema };
