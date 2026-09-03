import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './convert-mail-to-markdown.ts';

// The command's own tests live in commands.test.ts; this file covers the
// behaviours mutation testing showed were unpinned — chiefly the ways a single
// inline image can fail to embed without taking the whole message down with it.

const ATTS = '/me/messages/msg-1/attachments?$select=id,name,contentType,size,isInline,microsoft.graph.fileAttachment/contentId';
const logoImg = '<img src="cid:logo@fabrikam">';

const mail = (over: Record<string, unknown> = {}): Record<string, unknown> => ({
  subject: 'Q3',
  body: { contentType: 'html', content: `<p>text</p>${logoImg}` },
  hasAttachments: true,
  ...over,
});

const inlineLogo = (over: Record<string, unknown> = {}): Record<string, unknown> => ({
  id: 'att-1',
  name: 'logo.png',
  contentType: 'image/png',
  size: 120,
  isInline: true,
  contentId: 'logo@fabrikam',
  ...over,
});

const graphWith = (byPath: Record<string, Awaited<ReturnType<GraphClient['get']>>>): { graph: GraphClient; gets: string[] } => {
  const gets: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      gets.push(path);
      return byPath[path] ?? ok({});
    },
  });
  return { graph, gets };
};

const textOf = (value: unknown): string => (value as { text: string }).text;

// turndown escapes `[` / `]` on the way to markdown, so the placeholder the
// command writes as `[inline image: x]` reads as `\[inline image: x\]` here.
const PLACEHOLDER = String.raw`\[inline image: logo.png\]`;

describe('rendering an email whose inline image cannot be embedded', () => {
  it('leaves the image as a readable placeholder when Graph lists it without an id to fetch it by', async () => {
    const { graph, gets } = graphWith({
      '/me/messages/msg-1': ok(mail()),
      [ATTS]: ok({ value: [inlineLogo({ id: undefined })] }),
    });

    const result = await execute(graph, { messageId: 'msg-1', inlineImages: 'true' });

    expect(result.ok).toBe(true);
    if (result.ok) expect(textOf(result.value)).toContain(PLACEHOLDER);
    // There is no id to build a fetch URL from, so no bytes call is attempted.
    expect(gets.some((p) => p.includes('/attachments/'))).toBe(false);
  });

  it('keeps rendering the rest of the message when one image’s bytes fail to fetch', async () => {
    const { graph } = graphWith({
      '/me/messages/msg-1': ok(mail()),
      [ATTS]: ok({ value: [inlineLogo()] }),
      '/me/messages/msg-1/attachments/att-1': err({ type: 'api_error', status: 503, message: 'ServiceUnavailable' }),
    });

    const result = await execute(graph, { messageId: 'msg-1', inlineImages: 'true' });

    // One image failing must not fail the message: the text is what was asked for.
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(textOf(result.value)).toContain('text');
      expect(textOf(result.value)).toContain(PLACEHOLDER);
    }
  });

  it('falls back to a placeholder when Graph returns the attachment with no bytes in it', async () => {
    const { graph } = graphWith({
      '/me/messages/msg-1': ok(mail()),
      [ATTS]: ok({ value: [inlineLogo()] }),
      '/me/messages/msg-1/attachments/att-1': ok({ contentBytes: '' }),
    });

    const result = await execute(graph, { messageId: 'msg-1', inlineImages: 'true' });

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(textOf(result.value)).toContain(PLACEHOLDER);
      expect(textOf(result.value)).not.toContain('data:image/png;base64,');
    }
  });
});

describe('deciding which attachments count as embeddable inline images', () => {
  // One case per conjunct of the predicate. Each of these IS a file attachment
  // and must be listed as one rather than embedded into the body.
  const notInline = [
    { label: 'an attachment the sender added as a file rather than inline', over: { isInline: false, name: 'report.png' } },
    { label: 'an inline attachment that is not an image', over: { contentType: 'application/pdf', name: 'contract.pdf' } },
    { label: 'an inline image with no content id to match against the body', over: { contentId: undefined, name: 'orphan.png' } },
  ];

  it.each(notInline)('lists $label in the attachments table instead of embedding it', async ({ over }) => {
    const { graph } = graphWith({
      '/me/messages/msg-1': ok(mail()),
      [ATTS]: ok({ value: [inlineLogo(over)] }),
    });

    const result = await execute(graph, { messageId: 'msg-1', inlineImages: 'true' });

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(textOf(result.value)).toContain('**Attachments:**');
      expect(textOf(result.value)).not.toContain('data:');
    }
  });

  it('embeds an inline image whose content type is spelled in capitals, since the wire casing is not the sender’s choice', async () => {
    const { graph } = graphWith({
      '/me/messages/msg-1': ok(mail()),
      [ATTS]: ok({ value: [inlineLogo({ contentType: 'IMAGE/PNG' })] }),
      '/me/messages/msg-1/attachments/att-1': ok({ contentBytes: 'QUJD' }),
    });

    const result = await execute(graph, { messageId: 'msg-1', inlineImages: 'true' });

    if (result.ok) expect(textOf(result.value)).toContain('data:IMAGE/PNG;base64,QUJD');
  });
});

describe('rendering an email whose surrounding metadata is incomplete', () => {
  it('returns the body with a note naming the failure when the attachment list cannot be read', async () => {
    const { graph } = graphWith({
      '/me/messages/msg-1': ok(mail()),
      [ATTS]: err({ type: 'api_error', status: 403, message: 'Forbidden' }),
    });

    const result = await execute(graph, { messageId: 'msg-1' });

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(textOf(result.value)).toContain('text');
      expect((result.value as { note: string }).note).toContain('attachments-list fetch failed');
    }
  });

  it('omits the To: line entirely rather than printing an empty one when a message has no recipients', async () => {
    const { graph } = graphWith({
      '/me/messages/msg-1': ok(mail({ toRecipients: [], hasAttachments: false })),
    });

    const result = await execute(graph, { messageId: 'msg-1' });

    if (result.ok) expect(textOf(result.value)).not.toContain('**To:**');
  });

  it('names a sender who has an address but no display name by their address alone', async () => {
    const { graph } = graphWith({
      '/me/messages/msg-1': ok(mail({ from: { emailAddress: { address: 'robin.chen@fabrikam.com' } }, hasAttachments: false })),
    });

    const result = await execute(graph, { messageId: 'msg-1' });

    if (result.ok) expect(textOf(result.value)).toContain('**From:** robin.chen@fabrikam.com');
  });

  it('passes a failed message read through untouched', async () => {
    const { graph } = graphWith({ '/me/messages/msg-1': err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound' }) });

    expect(await execute(graph, { messageId: 'msg-1' })).toEqual(err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound' }));
  });
});

// Graph reports `hasAttachments: false` for a message whose only attachments are
// inline, so gating the attachments list on the flag hid exactly the images the
// body referenced, the trap get-mail-signature sidestepped on 2026-07-19.
describe('rendering an email whose only attachments are inline images, which Graph reports as hasAttachments: false', () => {
  it('still lists the inline image by name and id so the caller can fetch it, and names its placeholder after it', async () => {
    const { graph, gets } = graphWith({
      '/me/messages/msg-1': ok(mail({ hasAttachments: false })),
      [ATTS]: ok({ value: [inlineLogo()] }),
    });

    const result = await execute(graph, { messageId: 'msg-1' });

    expect(gets).toEqual(['/me/messages/msg-1', ATTS]);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(textOf(result.value)).toContain(PLACEHOLDER);
      expect(textOf(result.value)).toContain('- logo.png (120 B, image/png, id: att-1)');
    }
  });

  it('embeds the inline image under --inline-images true instead of silently skipping it', async () => {
    const { graph, gets } = graphWith({
      '/me/messages/msg-1': ok(mail({ hasAttachments: false })),
      [ATTS]: ok({ value: [inlineLogo()] }),
      '/me/messages/msg-1/attachments/att-1': ok({ contentBytes: 'QUJD' }),
    });

    const result = await execute(graph, { messageId: 'msg-1', inlineImages: 'true' });

    expect(gets).toEqual(['/me/messages/msg-1', ATTS, '/me/messages/msg-1/attachments/att-1']);
    if (result.ok) expect(textOf(result.value)).toContain('data:image/png;base64,QUJD');
  });

  it('still makes no attachments call for a message with no attachments and no cid: reference', async () => {
    const { graph, gets } = graphWith({
      '/me/messages/msg-1': ok(mail({ hasAttachments: false, body: { contentType: 'html', content: '<p>text only</p>' } })),
    });

    const result = await execute(graph, { messageId: 'msg-1' });

    expect(result.ok).toBe(true);
    expect(gets).toEqual(['/me/messages/msg-1']);
  });
});
