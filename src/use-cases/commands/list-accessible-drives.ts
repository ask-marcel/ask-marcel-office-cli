import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

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

type Source = 'channel' | 'joinedTeam' | 'memberOfGroup' | 'personal' | 'sharedWithMe';
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

  const value = [...drives.entries()]
    .map(([id, d]) => ({ id, name: d.name, driveType: d.driveType, webUrl: d.webUrl, sources: [...d.sources].sort(), ...(d.groupId !== null ? { groupId: d.groupId } : {}) }))
    .sort((a, b) => a.id.localeCompare(b.id));

  if (value.length === 0 && partialErrors.length > 0) return err(partialErrors[0]?.error ?? { type: 'api_error', status: 500, message: 'all drive-discovery sub-requests failed' });

  return ok({ value, count: value.length, ...(truncated ? { truncated: true } : {}), ...(partialErrors.length > 0 ? { partialErrors } : {}) });
};

const meta: CommandMeta = {
  summary:
    'Enumerate every drive (document library) the signed-in user can reach — personal OneDrive(s), Teams libraries, SharePoint M365-group sites, the drives behind files shared with the user, AND private/shared Teams channel sites — by unioning `/me/drives`, `/me/joinedTeams`, `/me/memberOf` (Unified groups, each resolved to `/groups/{id}/drive`), `/me/drive/sharedWithMe`, and per-team `/teams/{id}/channels` → `/channels/{ch}/filesFolder` (only private/shared channels, whose files live in their own site rather than the team default drive). Unlike `search-sharepoint-sites-by-name` (which relies on the tenant search index and misses direct-link-only sites), the sharedWithMe + channel vectors surface drives the search index never returns. Each drive is tagged with the `sources[]` that found it (a drive can have several). Per-vector and per-group failures do not fail the command — they appear in `partialErrors[]`; a group/channel with no provisioned drive is dropped silently. Fans out one `/groups/{id}/drive` + one `/teams/{id}/channels` call per joined team + member group, plus a `filesFolder` call per private/shared channel (all capped by `--max-groups`, default 100; raise carefully — large memberships can hit 429 throttling). `/me/followedSites` is not used — it 403s on this token.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drives + /me/joinedTeams + /me/memberOf + /me/drive/sharedWithMe + per-group /groups/<id>/drive + per-team /teams/<id>/channels/<ch>/filesFolder',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-list',
  options: [
    {
      name: 'max-groups',
      key: 'maxGroups',
      required: false,
      description:
        'Safety cap on each fan-out (positive integer; default 100): the per-group `/groups/{id}/drive` calls, the per-team `/teams/{id}/channels` enumeration, the per-private/shared-channel `filesFolder` lookups, and the `/drives/{id}` enrichment for shared- and channel-only drives. A user in hundreds of teams/groups would otherwise issue hundreds of parallel requests (429 risk). When any cap is hit the response carries `truncated: true`.',
      argumentHint: { kind: 'magicValue', values: ['100'] },
    },
  ],
  example: 'ask-marcel list-accessible-drives --output json',
  responseShape:
    '`{ value: [{ id, name, driveType, webUrl, sources: ["channel"|"joinedTeam"|"memberOfGroup"|"personal"|"sharedWithMe"], groupId? }], count, truncated?: true, partialErrors?: [{ source, error }] }`. `value[]` is deduped by drive `id` and sorted by id; `sources[]` lists every vector that surfaced the drive (`channel` = a private/shared Teams channel files folder); `groupId` is present only for Teams/group drives. `truncated: true` means a `--max-groups` cap was hit — raise it to see more. `partialErrors[]` (when present) names each vector, group, or channel whose sub-call failed.',
};

export { execute, meta, schema };
