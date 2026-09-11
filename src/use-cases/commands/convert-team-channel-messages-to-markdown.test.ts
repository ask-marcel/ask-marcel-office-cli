import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['convert-team-channel-messages-to-markdown'];
if (!command) throw new Error('convert-team-channel-messages-to-markdown is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const BASE = `/teams/tm1/channels/${CHANNEL}`;
const NAME = `${BASE}?$select=displayName`;
const PAGE = `${BASE}/messages?$top=50&$expand=replies`;
const SINCE = '2026-09-01T00:00:00Z';
const DELTA = `${BASE}/messages/delta?$filter=lastModifiedDateTime%20gt%20${SINCE}&$top=50&$expand=replies`;
const params = { teamId: 'tm1', channelId: CHANNEL };

const alex = { user: { id: 'u1', displayName: 'Alex Kim' } };
const post = (id: string, when: string, html: string, extra: Record<string, unknown> = {}): Record<string, unknown> => ({
  id,
  replyToId: null,
  messageType: 'message',
  createdDateTime: when,
  from: alex,
  body: { contentType: 'html', content: html },
  ...extra,
});
const event = {
  id: 'e1',
  messageType: 'unknownFutureValue',
  createdDateTime: '2026-09-07T00:00:00Z',
  eventDetail: { '@odata.type': '#microsoft.graph.membersAddedEventMessageDetail' },
};

type Reply = Awaited<ReturnType<GraphClient['get']>>;
type Envelope = { readonly contentType: string; readonly size: number; readonly text: string; readonly note?: string };

const render = async (byPath: Record<string, Reply>, extra: Record<string, string> = {}): Promise<{ result: Awaited<ReturnType<typeof command.execute>>; gets: string[] }> => {
  const gets: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      gets.push(path);
      return byPath[path] ?? err({ type: 'api_error', status: 404, message: `unexpected path ${path}` });
    },
  });
  return { result: await command.execute(graph, { ...params, ...extra }), gets };
};

describe('rendering a channel as a transcript', () => {
  it('without --since, reads the channel name and one page of posts with replies expanded, oldest post first', async () => {
    const { result, gets } = await render({
      [NAME]: ok({ displayName: 'General' }),
      [PAGE]: ok({ value: [post('2', '2026-09-09T09:00:00Z', '<p>Second</p>'), event, post('1', '2026-09-08T09:00:00Z', '<p>First</p>')] }),
    });
    expect(gets).toEqual([NAME, PAGE]);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.contentType).toBe('text/markdown');
    expect(envelope.text).toBe(
      [
        '# Transcript of General',
        '',
        '_Times are UTC, oldest post first._',
        '',
        '### 2026-09-08 09:00 · Alex Kim',
        '',
        'First',
        '',
        '### 2026-09-09 09:00 · Alex Kim',
        '',
        'Second',
      ].join('\n')
    );
    expect(envelope.note).toBe('2 posts on the newest page; 1 system event omitted');
    expect(envelope.size).toBe(new TextEncoder().encode(envelope.text).byteLength);
  });

  it('keeps rendering with a generic title when the channel name cannot be read, and honours --top on the page', async () => {
    const { result, gets } = await render({ [`${BASE}/messages?$top=5&$expand=replies`]: ok({ value: [post('1', '2026-09-08T09:00:00Z', '<p>Only</p>')] }) }, { top: '5' });
    expect(gets[0]).toBe(NAME);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.text.startsWith('# Channel transcript\n')).toBe(true);
    expect(envelope.note).toBe('1 post on the newest page');
  });

  it('with --since, walks the delta pages until the deltaLink and reports how many pages it read', async () => {
    const next = `${BASE}/messages/delta?$skiptoken=p2`;
    const { result, gets } = await render(
      {
        [NAME]: ok({ displayName: 'General' }),
        [DELTA]: ok({ value: [post('1', '2026-09-08T09:00:00Z', '<p>First</p>')], '@odata.nextLink': `https://graph.microsoft.com/v1.0${next}` }),
        [next]: ok({
          value: [post('2', '2026-09-09T09:00:00Z', '<p>Second</p>'), event],
          '@odata.deltaLink': `https://graph.microsoft.com/v1.0${BASE}/messages/delta?$deltatoken=t`,
        }),
      },
      { since: SINCE }
    );
    expect(gets).toEqual([NAME, DELTA, next]);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.text).toContain('First');
    expect(envelope.text).toContain('Second');
    expect(envelope.note).toBe(`2 posts touched since ${SINCE}, read in 2 pages; 1 system event omitted`);
  });

  it('accepts a relative --since and resolves it to an ISO instant in the delta filter', async () => {
    const { gets } = await render({}, { since: '7d' });
    const delta = gets[1] ?? '';
    expect(delta.startsWith(`${BASE}/messages/delta?$filter=lastModifiedDateTime%20gt%2020`)).toBe(true);
    expect(delta).toContain('Z&$top=50&$expand=replies');
  });

  it('stops the walk at --max-pages and says the range was cut', async () => {
    const byPath: Record<string, Reply> = { [NAME]: ok({ displayName: 'General' }) };
    let path = DELTA;
    for (let page = 1; page <= 3; page += 1) {
      const next = `${BASE}/messages/delta?$skiptoken=p${page}`;
      byPath[path] = ok({ value: [post(`${page}`, `2026-09-0${page}T09:00:00Z`, `<p>Post ${page}</p>`)], '@odata.nextLink': `https://graph.microsoft.com/v1.0${next}` });
      path = next;
    }
    const { result, gets } = await render(byPath, { since: SINCE, maxPages: '2' });
    expect(gets.length).toBe(3);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.text).toContain('Post 2');
    expect(envelope.text).not.toContain('Post 3');
    expect(envelope.note).toBe(`2 posts touched since ${SINCE}, read in 2 pages; truncated at --max-pages 2, older posts remain: raise --max-pages or narrow --since`);
  });

  it('names the channel behind the `410 Gone: UnknownError` an unknown channel answers', async () => {
    const { result } = await render({ [PAGE]: err({ type: 'api_error', status: 410, message: 'Gone: UnknownError' }) });
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_not_found');
  });

  it('refuses a page above 50, a bad date, and a max-pages outside 1 to 50 before calling Graph', async () => {
    const rejected: ReadonlyArray<Record<string, string>> = [{ top: '51' }, { since: 'next tuesday-ish' }, { maxPages: '0' }, { maxPages: '51' }];
    for (const extra of rejected) {
      const { result, gets } = await render({}, extra);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('validation_error');
      expect(gets).toEqual([]);
    }
  });

  it('advertises team-id, channel-id, since, top, max-pages and inline-images, produces bytes, and names the delta route it walks', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'channel-id', 'since', 'top', 'max-pages', 'inline-images']);
    expect(command.meta.category).toBe('teams');
    expect(command.meta.producesBytes).toBe(true);
    expect(command.meta.pagination).toBeUndefined();
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/channels/{channel-id}/messages/delta?$filter=lastModifiedDateTime gt {since}&$expand=replies');
  });
});
