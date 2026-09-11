import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['convert-team-channel-messages-to-markdown'];
if (!command) throw new Error('convert-team-channel-messages-to-markdown is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const BASE = `/teams/tm1/channels/${CHANNEL}`;
const HOSTED = 'https://graph.microsoft.com/v1.0/teams/tm1/channels/c1/messages/1/hostedContents/h1/$value';
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

describe('the page walk boundaries', () => {
  it('without --since, never follows a nextLink on the newest page and never says the range was cut', async () => {
    const { result, gets } = await render({
      [PAGE]: ok({ value: [post('1', '2026-09-08T09:00:00Z', '<p>Only</p>')], '@odata.nextLink': `https://graph.microsoft.com/v1.0${BASE}/messages?$skiptoken=p2` }),
    });
    expect(gets).toEqual([NAME, PAGE]);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as Envelope).note).toBe('1 post on the newest page');
  });

  it('accepts a two-digit --max-pages', async () => {
    const { result, gets } = await render({ [DELTA]: ok({ value: [] }) }, { since: SINCE, maxPages: '25' });
    expect(result.ok).toBe(true);
    expect(gets).toEqual([NAME, DELTA]);
  });
});

describe('what the transcript does with images and odd channel answers', () => {
  it('embeds hosted images only with --inline-images true, fetching them through Graph', async () => {
    const withImage = post('1', '2026-09-08T09:00:00Z', `<p>chart</p><img src="${HOSTED}" alt="chart">`);
    const binaries: string[] = [];
    const graph = (byPath: Record<string, Reply>): GraphClient =>
      fakeGraphClient({
        get: async (path) => byPath[path] ?? err({ type: 'api_error', status: 404, message: `unexpected path ${path}` }),
        getBinary: async (path) => {
          binaries.push(path);
          return ok({ contentType: 'image/png', size: 4, base64: 'AAAA' });
        },
      });
    const on = await command.execute(graph({ [NAME]: ok({ displayName: 'General' }), [PAGE]: ok({ value: [withImage] }) }), { ...params, inlineImages: 'true' });
    expect(binaries).toEqual(['/teams/tm1/channels/c1/messages/1/hostedContents/h1/$value']);
    if (on.ok) expect((on.value as Envelope).text).toContain('![chart](data:image/png;base64,AAAA)');
    const off = await command.execute(graph({ [NAME]: ok({ displayName: 'General' }), [PAGE]: ok({ value: [withImage] }) }), { ...params, inlineImages: 'false' });
    expect(binaries.length).toBe(1);
    if (off.ok) expect((off.value as Envelope).text).toContain('[image: hosted in Teams, pass --inline-images true to embed]');
  });

  it('uses the generic title when the channel answers a non-string name, and treats a page without a value list as empty', async () => {
    const { result } = await render({ [NAME]: ok({ displayName: 42 }), [PAGE]: ok({}) });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.text.startsWith('# Channel transcript\n')).toBe(true);
    expect(envelope.note).toBe('0 posts on the newest page');
  });

  it('names the accepted range when --max-pages is out of it', async () => {
    const { result } = await render({}, { since: SINCE, maxPages: '51' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('from 1 to 50');
  });
});
