import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { filterOutArchivedSites } from './filter-archived-sites.ts';

// A graph fake whose GET is routed by the probed site id, recording every probed path.
const graphProbing = (byId: (id: string) => Result<unknown, GraphError>, probed: Array<string>): GraphClient =>
  fakeGraphClient({
    get: async (path) => {
      probed.push(path);
      const id = /\/sites\/([^?]+)\?/.exec(path)?.[1] ?? '';
      return byId(id);
    },
  });

const active = ok({ id: 's', webUrl: 'https://contoso.sharepoint.com/sites/Team' });
const withStatus = (archiveStatus: string): Result<unknown, GraphError> => ok({ siteCollection: { archivalDetails: { archiveStatus } } });
const ids = (sites: ReadonlyArray<unknown>): Array<string> => sites.map((s) => (s as { id: string }).id);

describe('filterOutArchivedSites', () => {
  it('excludes a site whose metadata reports it fully archived, keeping the active one', async () => {
    const probed: Array<string> = [];
    const graph = graphProbing((id) => (id === 's2' ? withStatus('fullyArchived') : active), probed);
    const result = await filterOutArchivedSites(graph, [{ id: 's1' }, { id: 's2' }, { id: 's3' }]);
    expect(ids(result.value)).toEqual(['s1', 's3']);
    expect(result.archivedExcluded).toBe(1);
    expect(probed).toContain('/sites/s2?$select=id,webUrl,siteCollection&$expand=drive($select=quota)');
  });

  it('excludes recently-archived and reactivating sites as well', async () => {
    const byId = (id: string): Result<unknown, GraphError> => {
      if (id === 'a') return withStatus('recentlyArchived');
      if (id === 'b') return withStatus('reactivating');
      return active;
    };
    const graph = graphProbing(byId, []);
    const result = await filterOutArchivedSites(graph, [{ id: 'a' }, { id: 'b' }, { id: 'c' }]);
    expect(ids(result.value)).toEqual(['c']);
    expect(result.archivedExcluded).toBe(2);
  });

  it('treats a locked site (HTTP 423 resourceLocked) as archived and excludes it — the auto-archived OneDrive case', async () => {
    const graph = graphProbing(
      (id) => (id === 'locked' ? err({ type: 'api_error', status: 423, message: 'resourceLocked: Access to this site has been blocked.', code: 'resourceLocked' }) : active),
      []
    );
    const result = await filterOutArchivedSites(graph, [{ id: 'locked' }, { id: 'live' }]);
    expect(ids(result.value)).toEqual(['live']);
    expect(result.archivedExcluded).toBe(1);
    expect(result.probeErrors).toBe(0);
  });

  it('keeps a site whose probe fails for an unrelated reason and counts it as a probe error', async () => {
    const graph = graphProbing((id) => (id === 'flaky' ? err({ type: 'api_error', status: 500, message: 'boom' }) : active), []);
    const result = await filterOutArchivedSites(graph, [{ id: 'flaky' }, { id: 'ok' }]);
    expect(ids(result.value)).toEqual(['flaky', 'ok']);
    expect(result.probeErrors).toBe(1);
    expect(result.archivedExcluded).toBe(0);
  });

  it('passes active sites through unchanged with nothing excluded', async () => {
    const result = await filterOutArchivedSites(
      graphProbing(() => active, []),
      [{ id: 's1' }, { id: 's2' }]
    );
    expect(ids(result.value)).toEqual(['s1', 's2']);
    expect(result.archivedExcluded).toBe(0);
    expect(result.probeErrors).toBe(0);
    expect(result.probeTruncated).toBe(false);
  });

  it('attaches per-site size from the expanded drive quota onto a kept site', async () => {
    const probeResp = ok({ id: 's1', webUrl: 'https://contoso.sharepoint.com/sites/Team', drive: { quota: { used: 9000 } } });
    const result = await filterOutArchivedSites(
      graphProbing(() => probeResp, []),
      [{ id: 's1', webUrl: 'https://contoso.sharepoint.com/sites/Team' }]
    );
    expect(result.value).toEqual([{ id: 's1', webUrl: 'https://contoso.sharepoint.com/sites/Team', size: 9000 }]);
  });

  it('keeps a site without a size field when the probe returns no drive quota', async () => {
    const result = await filterOutArchivedSites(
      graphProbing(() => active, []),
      [{ id: 's1', webUrl: 'https://contoso.sharepoint.com/sites/Team' }]
    );
    expect((result.value[0] as { size?: number }).size).toBeUndefined();
  });

  it('keeps a null hit without probing or throwing', async () => {
    const probed: Array<string> = [];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, probed),
      [null, { id: 's1' }]
    );
    expect(result.value.length).toBe(2);
    expect(result.archivedExcluded).toBe(0);
    expect(probed).toEqual(['/sites/s1?$select=id,webUrl,siteCollection&$expand=drive($select=quota)']); // null was never probed
  });

  it('keeps a hit that has no id without probing it', async () => {
    const probed: Array<string> = [];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, probed),
      [{ name: 'no id' }, { id: 's1' }]
    );
    expect(result.value.length).toBe(2);
    expect(probed).toEqual(['/sites/s1?$select=id,webUrl,siteCollection&$expand=drive($select=quota)']);
  });

  it('probes every site across multiple concurrency chunks', async () => {
    const probed: Array<string> = [];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, probed),
      [{ id: 'a' }, { id: 'b' }, { id: 'c' }],
      { chunkSize: 2 }
    );
    expect(probed.length).toBe(3);
    expect(ids(result.value)).toEqual(['a', 'b', 'c']);
  });

  it('stops probing past the cap, keeps the remainder unprobed, and flags the result truncated', async () => {
    const probed: Array<string> = [];
    const result = await filterOutArchivedSites(
      graphProbing(() => withStatus('fullyArchived'), probed),
      [{ id: 's1' }, { id: 's2' }],
      { probeMax: 1 }
    );
    expect(probed).toEqual(['/sites/s1?$select=id,webUrl,siteCollection&$expand=drive($select=quota)']); // s2 never probed
    expect(result.probeTruncated).toBe(true);
    expect(ids(result.value)).toEqual(['s2']); // s1 dropped (archived), s2 kept unprobed
    expect(result.archivedExcluded).toBe(1);
  });

  it('is not truncated when the site count exactly equals the cap', async () => {
    const result = await filterOutArchivedSites(
      graphProbing(() => active, []),
      [{ id: 's1' }, { id: 's2' }],
      { probeMax: 2 }
    );
    expect(result.probeTruncated).toBe(false);
  });
});

// Each archived signal in isolation: the others must NOT also fire, so the deciding
// branch is exercised on its own (status / code / message short-circuit independently).
const verdictFor = async (probeResult: Result<unknown, GraphError>): Promise<{ kept: boolean; archived: number; errors: number }> => {
  const r = await filterOutArchivedSites(
    graphProbing(() => probeResult, []),
    [{ id: 'x' }]
  );
  return { kept: r.value.length === 1, archived: r.archivedExcluded, errors: r.probeErrors };
};

describe('filterOutArchivedSites — archive signal detection', () => {
  it('drops on HTTP 423 alone (no archive code, non-matching message)', async () => {
    expect(await verdictFor(err({ type: 'api_error', status: 423, message: 'temporarily unavailable' }))).toEqual({ kept: false, archived: 1, errors: 0 });
  });

  it('drops on a resourceLocked code alone, case-insensitively (status not 423, message non-matching)', async () => {
    expect(await verdictFor(err({ type: 'api_error', status: 403, message: 'denied', code: 'resourceLocked' }))).toEqual({ kept: false, archived: 1, errors: 0 });
  });

  it('drops on a siteArchived code alone, case-insensitively', async () => {
    expect(await verdictFor(err({ type: 'api_error', status: 404, message: 'gone', code: 'SiteArchived' }))).toEqual({ kept: false, archived: 1, errors: 0 });
  });

  it('drops on an "archived" message alone (status not 423, code not an archive code)', async () => {
    expect(await verdictFor(err({ type: 'api_error', status: 500, message: 'this site is archived', code: 'generalException' }))).toEqual({ kept: false, archived: 1, errors: 0 });
  });

  it('drops on a "has been blocked" message alone', async () => {
    expect(await verdictFor(err({ type: 'api_error', status: 500, message: 'access has been blocked', code: 'generalException' }))).toEqual({
      kept: false,
      archived: 1,
      errors: 0,
    });
  });

  it('drops on a sharepointerror message alone', async () => {
    expect(await verdictFor(err({ type: 'api_error', status: 500, message: 'redirected to sharepointerror.aspx', code: 'generalException' }))).toEqual({
      kept: false,
      archived: 1,
      errors: 0,
    });
  });

  it('drops on an archive message carried by a non-api_error (no status field at all)', async () => {
    expect(await verdictFor(err({ type: 'network_error', message: 'the site is archived' }))).toEqual({ kept: false, archived: 1, errors: 0 });
  });

  it('keeps a site when none of status, code or message signals archival', async () => {
    expect(await verdictFor(err({ type: 'api_error', status: 500, message: 'internal error', code: 'generalException' }))).toEqual({ kept: true, archived: 0, errors: 1 });
  });

  it('keeps a site whose metadata probe returns null (no throw on the null guard)', async () => {
    expect(await verdictFor(ok(null))).toEqual({ kept: true, archived: 0, errors: 0 });
  });

  it('keeps a site whose archiveStatus is present but not a string', async () => {
    expect(await verdictFor(ok({ siteCollection: { archivalDetails: { archiveStatus: 123 } } }))).toEqual({ kept: true, archived: 0, errors: 0 });
  });

  it('keeps a site whose siteCollection carries no archivalDetails', async () => {
    expect(await verdictFor(ok({ siteCollection: { hostname: 'contoso.sharepoint.com' } }))).toEqual({ kept: true, archived: 0, errors: 0 });
  });
});

describe('filterOutArchivedSites — non-navigable + not-found exclusion', () => {
  it('excludes a SharePoint Add-in app-domain site by its webUrl WITHOUT probing it', async () => {
    const probed: Array<string> = [];
    const sites = [
      { id: 'addin', webUrl: 'https://contoso-d870846bb167ac.sharepoint.com/sites/Apps/_layouts/15/AddinDeprecationAnnoucement.aspx' },
      { id: 'live', webUrl: 'https://contoso.sharepoint.com/sites/Team' },
    ];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, probed),
      sites
    );
    expect(ids(result.value)).toEqual(['live']);
    expect(result.nonNavigableExcluded).toBe(1);
    expect(probed).not.toContain('/sites/addin?$select=id,webUrl,siteCollection&$expand=drive($select=quota)');
  });

  it('excludes a SharePoint Embedded /contentstorage/ container by its webUrl', async () => {
    const sites = [
      { id: 'cs', webUrl: 'https://contoso.sharepoint.com/contentstorage/coJsE0OdIkqu2uEOCncHOTHn9Wdld5BMqhxRcKelXVg' },
      { id: 'live', webUrl: 'https://contoso.sharepoint.com/sites/Team' },
    ];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, []),
      sites
    );
    expect(ids(result.value)).toEqual(['live']);
    expect(result.nonNavigableExcluded).toBe(1);
  });

  it('excludes a bare /_layouts/ system URL on an otherwise-normal host', async () => {
    const sites = [
      { id: 'sys', webUrl: 'https://contoso.sharepoint.com/sites/X/_layouts/15/viewlsts.aspx' },
      { id: 'live', webUrl: 'https://contoso.sharepoint.com/sites/Team' },
    ];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, []),
      sites
    );
    expect(ids(result.value)).toEqual(['live']);
    expect(result.nonNavigableExcluded).toBe(1);
  });

  it('excludes a site whose probe returns HTTP 404, counted separately from archived', async () => {
    const graph = graphProbing((id) => (id === 'gone' ? err({ type: 'api_error', status: 404, message: 'itemNotFound: site not found' }) : active), []);
    const sites = [
      { id: 'gone', webUrl: 'https://contoso.sharepoint.com/sites/Gone' },
      { id: 'live', webUrl: 'https://contoso.sharepoint.com/sites/Team' },
    ];
    const result = await filterOutArchivedSites(graph, sites);
    expect(ids(result.value)).toEqual(['live']);
    expect(result.notFoundExcluded).toBe(1);
    expect(result.archivedExcluded).toBe(0);
    expect(result.probeErrors).toBe(0);
  });

  it('keeps an active personal OneDrive that probes ok — OneDrives are not dropped by URL shape', async () => {
    const probed: Array<string> = [];
    const sites = [{ id: 'od', webUrl: 'https://contoso-my.sharepoint.com/personal/b_departed_user_hk_fabrikam_com' }];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, probed),
      sites
    );
    expect(ids(result.value)).toEqual(['od']);
    expect(result.nonNavigableExcluded).toBe(0);
    expect(result.notFoundExcluded).toBe(0);
    expect(probed).toContain('/sites/od?$select=id,webUrl,siteCollection&$expand=drive($select=quota)');
  });

  it('excludes an add-in app-domain host even with a normal path (no /_layouts/ or /contentstorage/ marker)', async () => {
    const probed: Array<string> = [];
    const sites = [{ id: 'addin2', webUrl: 'https://contoso-18dd384baca361.sharepoint.com/sites/Apps' }];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, probed),
      sites
    );
    expect(ids(result.value)).toEqual([]);
    expect(result.nonNavigableExcluded).toBe(1);
    expect(probed).not.toContain('/sites/addin2?$select=id,webUrl,siteCollection&$expand=drive($select=quota)');
  });

  it('keeps a host whose hyphen-suffix is too short to be an app domain (boundary on the hex-run length)', async () => {
    const probed: Array<string> = [];
    const sites = [{ id: 'short', webUrl: 'https://contoso-a.sharepoint.com/sites/X' }];
    const result = await filterOutArchivedSites(
      graphProbing(() => active, probed),
      sites
    );
    expect(ids(result.value)).toEqual(['short']);
    expect(result.nonNavigableExcluded).toBe(0);
    expect(probed).toContain('/sites/short?$select=id,webUrl,siteCollection&$expand=drive($select=quota)');
  });
});
