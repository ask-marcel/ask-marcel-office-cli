import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';
import { buildShareToken } from './sharepoint-link-extractor.ts';

const DENIED: GraphError = { type: 'api_error', status: 403, message: 'accessDenied: Access denied', code: 'accessDenied' };
const MISSING: GraphError = { type: 'api_error', status: 404, message: 'itemNotFound: The resource could not be found.', code: 'itemNotFound' };

const LINKS: ReadonlyArray<readonly [string, Result<unknown, GraphError>]> = [
  ['https://contoso-my.sharepoint.com/personal/robin_chen_contoso_com/Documents/Lifecycle.pptx', err(DENIED)],
  ['https://contoso-my.sharepoint.com/:p:/g/personal/alex_kim_contoso_com/EXyz', err(DENIED)],
  ['https://contoso.sharepoint.com/:w:/r/sites/Operations/Shared%20Documents/SOP.docx?d=w1', err(DENIED)],
  ['https://contoso.sharepoint.com/:x:/s/Finance/EabcDEF', err(DENIED)],
  ['https://contoso.sharepoint.com/:x:/t/Atlas/Eghi', err(DENIED)],
  ['https://contoso.sharepoint.com/teams/Atlas/Shared%20Documents/plan.xlsx', err(MISSING)],
  ['https://contoso.sharepoint.com/Shared%20Documents/root.docx', err(DENIED)],
  [
    'https://contoso.sharepoint.com/sites/Operations/Shared%20Documents/open.docx',
    ok({ id: 'i1', name: 'open.docx', webUrl: 'https://contoso.sharepoint.com/w', parentReference: { driveId: 'd1' } }),
  ],
  ['https://contoso-my.sharepoint.com/:fl:/g/personal/jordan_avery_contoso_com/EloopA', err(DENIED)],
  ['https://contoso.sharepoint.com/:fl:/r/sites/Wiki/SitePages/Access.aspx', err(DENIED)],
  ['https://contoso.sharepoint.com/:fl:/s/Board/Eboard', err({ type: 'network_error', message: 'socket hang up' })],
];

const ASK = 'open the link in a browser to send an access request.';

describe('a SharePoint link the user cannot open', () => {
  it('names where the link points, read from the URL, and how to ask for access when Graph refuses it', async () => {
    const byToken = new Map(LINKS.map(([url, answer]) => [buildShareToken(url), answer]));
    const graph = fakeGraphClient({
      get: async (path: string) => {
        if (path.startsWith('/me/messages/')) return ok({ subject: 'SOPs', body: { content: LINKS.map(([url]) => `<a href="${url}">x</a>`).join(' ') } });
        return byToken.get(path.replace('/shares/', '').replace('/driveItem', '')) ?? err(MISSING);
      },
    });
    const command = commands['extract-sharepoint-links-in-mail'];
    if (!command) throw new Error('extract-sharepoint-links-in-mail is not registered');
    const result = await command.execute(graph, { messageId: 'm1' });
    if (!result.ok) throw new Error(result.error.message);
    const links = (result.value as { links: ReadonlyArray<Record<string, unknown>> }).links;
    const onedrive = (owner: string): string =>
      `It sits in the OneDrive of ${owner} (a sign-in name with . and @ written as _), which has not been shared with you: ask the owner for access, or ${ASK}`;
    const site = (name: string): string => `It sits on the SharePoint site ${name}, which you cannot open: ask a site owner for access, or ${ASK}`;
    expect(links.map(({ location, hint }) => ({ location, hint }))).toEqual([
      { location: { kind: 'onedrive', owner: 'robin_chen_contoso_com' }, hint: onedrive('robin_chen_contoso_com') },
      { location: { kind: 'onedrive', owner: 'alex_kim_contoso_com' }, hint: onedrive('alex_kim_contoso_com') },
      { location: { kind: 'site', site: 'sites/Operations' }, hint: site('sites/Operations') },
      { location: { kind: 'site', site: 'sites/Finance' }, hint: site('sites/Finance') },
      { location: { kind: 'site', site: 'teams/Atlas' }, hint: site('teams/Atlas') },
      { location: { kind: 'site', site: 'teams/Atlas' }, hint: undefined },
      { location: undefined, hint: `Ask the file's owner for access, or ${ASK}` },
      { location: undefined, hint: undefined },
      { location: { kind: 'onedrive', owner: 'jordan_avery_contoso_com' }, hint: onedrive('jordan_avery_contoso_com') },
      { location: { kind: 'site', site: 'sites/Wiki' }, hint: site('sites/Wiki') },
      { location: { kind: 'site', site: 'sites/Board' }, hint: undefined },
    ]);
    expect(links[0]?.error).toBe('accessDenied: Access denied');
    expect(links[10]?.error).toBe('network_error: socket hang up');
    expect(Object.keys(links[6] ?? {})).toEqual(['url', 'error', 'hint']);
    expect(Object.keys(links[7] ?? {})).toEqual(['url', 'driveId', 'itemId', 'name', 'webUrl']);
  });
});
