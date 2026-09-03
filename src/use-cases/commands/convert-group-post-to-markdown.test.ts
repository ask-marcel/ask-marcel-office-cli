import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { lookupScopes } from './graph-scopes.ts';
import { commands } from './index.ts';

const command = commands['convert-group-post-to-markdown'];
if (!command) throw new Error('convert-group-post-to-markdown is not registered');

const POST = '/groups/g1/threads/t1/posts/p1';
const ATTS = `${POST}/attachments?$select=id,name,contentType,size,isInline,microsoft.graph.fileAttachment/contentId`;

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

const inlineLogo = { id: 'att-1', name: 'logo.png', contentType: 'image/png', size: 120, isInline: true, contentId: 'logo@contoso' };
const logoImg = '<img src="cid:logo@contoso">';

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

// turndown escapes `[` / `]` on the way to markdown.
const PLACEHOLDER = String.raw`\[inline image: logo.png\]`;

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

describe('the attachments of a group post', () => {
  it('lists a file attachment under the body with its size and id, and points at get-group-post for the bytes', async () => {
    const { result, gets } = await render({
      [POST]: ok(post({ hasAttachments: true })),
      [ATTS]: ok({ value: [{ id: 'att-2', name: 'report.pdf', contentType: 'application/pdf', size: 4_200_000, isInline: false, contentId: null }] }),
    });

    const { text } = envelopeOf(result);
    expect(gets).toEqual([POST, ATTS]);
    expect(text).toContain('**Attachments:**');
    expect(text).toContain('- report.pdf (4.2 MB, application/pdf, id: att-2)');
    expect(text).toContain('get-group-post');
    expect(text).not.toContain('get-mail-attachment');
  });

  it('embeds an inline image fetched from the post’s own attachments path when asked to, and leaves a file attachment listed rather than fetched', async () => {
    const { result, gets } = await render(
      {
        [POST]: ok(post({ hasAttachments: true, body: { contentType: 'html', content: `<p>Agenda</p>${logoImg}` } })),
        [ATTS]: ok({ value: [inlineLogo, { id: 'att-2', name: 'report.pdf', contentType: 'application/pdf', size: 4_200_000, isInline: false }] }),
        [`${POST}/attachments/att-1`]: ok({ contentBytes: 'QUJD' }),
      },
      { inlineImages: 'true' }
    );

    expect(gets).toEqual([POST, ATTS, `${POST}/attachments/att-1`]);
    const { text } = envelopeOf(result);
    expect(text).toContain('data:image/png;base64,QUJD');
    // An embedded image is not also listed; the file attachment still is.
    expect(text).toContain('- report.pdf');
    expect(text).not.toContain('- logo.png');
  });

  it('leaves an inline image as a readable placeholder by default, lists it so the caller knows it exists, and fetches no bytes', async () => {
    const { result, gets } = await render({
      [POST]: ok(post({ hasAttachments: true, body: { contentType: 'html', content: `<p>Agenda</p>${logoImg}` } })),
      [ATTS]: ok({ value: [inlineLogo] }),
    });

    expect(gets).toEqual([POST, ATTS]);
    const { text } = envelopeOf(result);
    expect(text).toContain(PLACEHOLDER);
    expect(text).toContain('- logo.png (120 B, image/png, id: att-1)');
    expect(text).not.toContain('data:image/png');
  });
});

describe('quoted replies inside a group post', () => {
  const quoted = '<p>My reply.</p><div id="divRplyFwdMsg"><b>From:</b> Alex Kim<br><b>Sent:</b> Friday</div><p>Original message text.</p>';

  it('strips the quoted reply chain by default and says so in the note', async () => {
    const { result } = await render({ [POST]: ok(post({ body: { contentType: 'html', content: quoted } })) });

    const envelope = envelopeOf(result);
    expect(envelope.text).toContain('My reply.');
    expect(envelope.text).not.toContain('Original message text');
    expect(envelope.note).toContain('quoted reply chain stripped');
  });

  it('keeps the quoted chain with --keep-quoted true', async () => {
    const { result } = await render({ [POST]: ok(post({ body: { contentType: 'html', content: quoted } })) }, { keepQuoted: 'true' });

    const envelope = envelopeOf(result);
    expect(envelope.text).toContain('Original message text');
    expect(envelope.note).toBeUndefined();
  });

  it('accepts an explicit false on both flags and behaves exactly as the default', async () => {
    const byPath = {
      [POST]: ok(post({ hasAttachments: true, body: { contentType: 'html', content: `${quoted}${logoImg}` } })),
      [ATTS]: ok({ value: [inlineLogo] }),
    };

    const explicit = await render(byPath, { keepQuoted: 'false', inlineImages: 'false' });
    const implicit = await render(byPath);

    expect(explicit.result).toEqual(implicit.result);
    expect(explicit.gets).toEqual([POST, ATTS]);
    expect(envelopeOf(explicit.result).note).toContain('quoted reply chain stripped');
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
