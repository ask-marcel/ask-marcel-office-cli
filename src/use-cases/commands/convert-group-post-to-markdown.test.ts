import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { lookupScopes } from './graph-scopes.ts';
import { commands } from './index.ts';

const command = commands['convert-group-post-to-markdown'];
if (!command) throw new Error('convert-group-post-to-markdown is not registered');

const POST = '/groups/g1/threads/t1/posts/p1';

// A post arrives FROM the group's own address; the person who wrote it is the
// `sender` (probed live 2026-09-03). All placeholder identities.
const support = { emailAddress: { name: 'Support', address: 'support@contoso.com' } };
const robin = { emailAddress: { name: 'Robin Chen', address: 'robin.chen@contoso.com' } };

const post = (over: Record<string, unknown> = {}): Record<string, unknown> => ({
  body: { contentType: 'html', content: '<p>Agenda for Monday</p>' },
  receivedDateTime: '2026-07-03T13:48:08Z',
  from: support,
  sender: robin,
  hasAttachments: false,
  ...over,
});

type Envelope = { readonly contentType: string; readonly size: number; readonly text: string; readonly note?: string };
type Reply = Awaited<ReturnType<GraphClient['get']>>;

const render = async (byPath: Record<string, Reply>, params: Record<string, string> = {}): Promise<{ result: Awaited<ReturnType<typeof command.execute>>; gets: string[] }> => {
  const gets: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      gets.push(path);
      return byPath[path] ?? ok({});
    },
  });
  const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1', ...params });
  return { result, gets };
};

const envelopeOf = (result: Awaited<ReturnType<typeof command.execute>>): Envelope => {
  if (!result.ok) throw new Error(`expected markdown, got ${result.error.type}: ${result.error.message}`);
  return result.value as Envelope;
};

describe('rendering a group post as markdown', () => {
  it('names the person who wrote the post on behalf of the group, dates it, and renders the body, in one call when nothing is attached', async () => {
    const { result, gets } = await render({ [POST]: ok(post()) });

    const envelope = envelopeOf(result);
    expect(gets).toEqual([POST]);
    expect(envelope.contentType).toBe('text/markdown');
    // The header block is exactly these two lines, then a blank line, then the body.
    expect(
      envelope.text.startsWith('**From:** Robin Chen <robin.chen@contoso.com> on behalf of Support <support@contoso.com>\n**Date:** 2026-07-03T13:48:08Z\n\nAgenda for Monday')
    ).toBe(true);
    expect(envelope.text).not.toContain('**Subject:**');
    expect(envelope.size).toBe(new TextEncoder().encode(envelope.text).byteLength);
    expect(envelope.note).toBeUndefined();
  });

  it('names the group alone when the post was written from its own address', async () => {
    const { result } = await render({ [POST]: ok(post({ sender: support })) });

    expect(envelopeOf(result).text).toContain('**From:** Support <support@contoso.com>');
    expect(envelopeOf(result).text).not.toContain('on behalf of');
  });

  it('names the sender alone when Graph returns no from address', async () => {
    const { result } = await render({ [POST]: ok(post({ from: undefined })) });

    expect(envelopeOf(result).text).toContain('**From:** Robin Chen <robin.chen@contoso.com>');
    expect(envelopeOf(result).text).not.toContain('on behalf of');
  });

  it('falls back to the from address when Graph returns no sender', async () => {
    const { result } = await render({ [POST]: ok(post({ sender: undefined })) });

    expect(envelopeOf(result).text).toContain('**From:** Support <support@contoso.com>');
  });

  it('renders a post with no author and no date as the body alone', async () => {
    const { result } = await render({ [POST]: ok(post({ from: undefined, sender: undefined, receivedDateTime: undefined })) });

    const { text } = envelopeOf(result);
    expect(text).not.toContain('**From:**');
    expect(text).not.toContain('**Date:**');
    expect(text).toContain('Agenda for Monday');
  });
});

describe('a group post that cannot be rendered', () => {
  it('passes a failed post read through untouched', async () => {
    const { result } = await render({ [POST]: err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound' }) });

    expect(result).toEqual(err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound' }));
  });

  it('refuses a call that names the group and post but not the thread, before any read', async () => {
    const gets: string[] = [];
    const graph = fakeGraphClient({
      get: async (path) => {
        gets.push(path);
        return ok(post());
      },
    });

    const result = await command.execute(graph, { groupId: 'g1', postId: 'p1' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(gets).toEqual([]);
  });

  it('is a mail command that produces bytes, on the group scope the token already carries', () => {
    expect(command.meta.category).toBe('mail');
    expect(command.meta.producesBytes).toBe(true);
    expect(lookupScopes('convert-group-post-to-markdown')).toEqual(['Group.Read.All']);
  });
});
