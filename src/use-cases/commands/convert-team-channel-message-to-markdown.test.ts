import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['convert-team-channel-message-to-markdown'];
if (!command) throw new Error('convert-team-channel-message-to-markdown is not registered');

const CHANNEL = '19:abc@thread.tacv2';
const MESSAGE = `/teams/tm1/channels/${CHANNEL}/messages/1700000000000`;
const REPLIES = `${MESSAGE}/replies?$top=50`;
const HOSTED = 'https://graph.microsoft.com/v1.0/teams/tm1/channels/c1/messages/1700000000000/hostedContents/h1/$value';
const params = { teamId: 'tm1', channelId: CHANNEL, messageId: '1700000000000' };

const alex = { user: { id: 'u1', displayName: 'Alex Kim' } };
const robin = { user: { id: 'u2', displayName: 'Robin Chen' } };
const root = { id: '1700000000000', messageType: 'message', createdDateTime: '2026-09-08T14:03:11Z', from: alex, body: { contentType: 'html', content: '<p>Budget review</p>' } };
const replyAt = (id: string, when: string, html: string): Record<string, unknown> => ({
  id,
  replyToId: '1700000000000',
  messageType: 'message',
  createdDateTime: when,
  from: robin,
  body: { contentType: 'html', content: html },
});

type Reply = Awaited<ReturnType<GraphClient['get']>>;
type Envelope = { readonly contentType: string; readonly size: number; readonly text: string; readonly note?: string };

const render = async (
  byPath: Record<string, Reply>,
  extra: Record<string, string> = {},
  graphOverrides: Partial<GraphClient> = {}
): Promise<{ result: Awaited<ReturnType<typeof command.execute>>; gets: string[] }> => {
  const gets: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      gets.push(path);
      return byPath[path] ?? err({ type: 'api_error', status: 404, message: `unexpected path ${path}` });
    },
    ...graphOverrides,
  });
  return { result: await command.execute(graph, { ...params, ...extra }), gets };
};

describe('rendering one channel post with its replies as markdown', () => {
  it('reads the post, then its replies fifty at a time, and renders the thread oldest reply first', async () => {
    const { result, gets } = await render({
      [MESSAGE]: ok(root),
      [REPLIES]: ok({ value: [replyAt('2', '2026-09-08T15:00:00Z', '<p>Second</p>'), replyAt('1', '2026-09-08T14:10:00Z', '<p>First</p>')] }),
    });
    expect(gets).toEqual([MESSAGE, REPLIES]);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.contentType).toBe('text/markdown');
    expect(envelope.text).toBe(
      ['### 2026-09-08 14:03 · Alex Kim', '', 'Budget review', '', '> **Robin Chen · 2026-09-08 14:10**', '> First', '', '> **Robin Chen · 2026-09-08 15:00**', '> Second'].join(
        '\n'
      )
    );
    expect(envelope.size).toBe(new TextEncoder().encode(envelope.text).byteLength);
    expect(envelope.note).toBeUndefined();
  });

  it('follows the replies nextLink page by page, and stops after ten pages with a note', async () => {
    const byPath: Record<string, Reply> = { [MESSAGE]: ok(root) };
    let path = REPLIES;
    for (let page = 1; page <= 12; page += 1) {
      const next = `${MESSAGE}/replies?$skiptoken=p${page}`;
      byPath[path] = ok({
        value: [replyAt(`r${page}`, `2026-09-08T14:${String(page).padStart(2, '0')}:00Z`, `<p>Reply ${page}</p>`)],
        '@odata.nextLink': `https://graph.microsoft.com/v1.0${next}`,
      });
      path = next;
    }
    const { result, gets } = await render(byPath);
    expect(gets.length).toBe(11);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.text).toContain('Reply 10');
    expect(envelope.text).not.toContain('Reply 11');
    expect(envelope.note).toBe('replies truncated after 10 pages (10 replies rendered); the thread has more');
  });

  it('renders the post alone with a note when the replies read fails', async () => {
    const { result } = await render({ [MESSAGE]: ok(root), [REPLIES]: err({ type: 'api_error', status: 503, message: 'ServiceUnavailable: try later' }) });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.text).toBe(['### 2026-09-08 14:03 · Alex Kim', '', 'Budget review'].join('\n'));
    expect(envelope.note).toBe('replies fetch failed (api_error: ServiceUnavailable: try later); the post is rendered without its replies');
  });

  it('embeds hosted images only when asked, fetching them through Graph', async () => {
    const withImage = { ...root, body: { contentType: 'html', content: `<p>chart</p><img src="${HOSTED}" alt="chart">` } };
    const binaries: string[] = [];
    const getBinary: GraphClient['getBinary'] = async (path) => {
      binaries.push(path);
      return ok({ contentType: 'image/png', size: 4, base64: 'AAAA' });
    };
    const off = await render({ [MESSAGE]: ok(withImage), [REPLIES]: ok({ value: [] }) }, {}, { getBinary });
    expect(binaries).toEqual([]);
    if (off.result.ok) expect((off.result.value as Envelope).text).toContain('[image: hosted in Teams, pass --inline-images true to embed]');
    const on = await render({ [MESSAGE]: ok(withImage), [REPLIES]: ok({ value: [] }) }, { inlineImages: 'true' }, { getBinary });
    expect(binaries).toEqual(['/teams/tm1/channels/c1/messages/1700000000000/hostedContents/h1/$value']);
    if (on.result.ok) expect((on.result.value as Envelope).text).toContain('![chart](data:image/png;base64,AAAA)');
  });

  it('names the message behind the `403 Forbidden: UnknownError` an unknown id answers', async () => {
    const { result } = await render({ [MESSAGE]: err({ type: 'api_error', status: 403, message: 'Forbidden: UnknownError' }) });
    expect(result.ok).toBe(false);
    if (result.ok || result.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(result.error.code).toBe('cli_rewrite_channel_message_not_found');
    expect(result.error.message).toContain('message-id: "1700000000000"');
  });

  it('refuses a missing message id before calling Graph', async () => {
    const { result, gets } = await render({}, { messageId: '' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(gets).toEqual([]);
  });

  it('advertises team-id, channel-id, message-id and inline-images, produces bytes, and lives with the teams commands', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['team-id', 'channel-id', 'message-id', 'inline-images']);
    expect(command.meta.category).toBe('teams');
    expect(command.meta.producesBytes).toBe(true);
    expect(command.meta.graphPathTemplate).toBe('/teams/{team-id}/channels/{channel-id}/messages/{message-id}');
  });
});

describe('the inline-images flag spelled out', () => {
  it('accepts --inline-images false and then fetches nothing', async () => {
    const binaries: string[] = [];
    const getBinary: GraphClient['getBinary'] = async (path) => {
      binaries.push(path);
      return ok({ contentType: 'image/png', size: 4, base64: 'AAAA' });
    };
    const { result } = await render(
      { [MESSAGE]: ok({ ...root, body: { contentType: 'html', content: `<img src="${HOSTED}">` } }), [REPLIES]: ok({ value: [] }) },
      { inlineImages: 'false' },
      { getBinary }
    );
    expect(result.ok).toBe(true);
    expect(binaries).toEqual([]);
  });
});

describe('a replies page without a value list', () => {
  it('renders the post alone, as an empty page would', async () => {
    const { result } = await render({ [MESSAGE]: ok(root), [REPLIES]: ok({}) });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as Envelope).text).toBe(['### 2026-09-08 14:03 · Alex Kim', '', 'Budget review'].join('\n'));
  });
});
