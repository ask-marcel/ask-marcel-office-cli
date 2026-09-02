import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './get-mail-signature.ts';

const SCAN_PATH = '/me/mailFolders/sentitems/messages?$top=10&$orderby=sentDateTime%20desc&$select=id,sentDateTime';
const SIGNATURE = '<div id="Signature"><div>Robin Chen</div><div>Fabrikam</div></div>';
const QUOTE_TAIL = '<div id="appendonsend"></div><div id="divRplyFwdMsg"><b>From:</b> Alex Kim<br><b>Sent:</b> Monday</div>';
const bodyOf = (id: string): string => `/me/messages/${id}?$select=body,sentDateTime,hasAttachments`;
const attsOf = (id: string): string => `/me/messages/${id}/attachments?${'$select=id,name,contentType,size,isInline,microsoft.graph.fileAttachment/contentId'}`;

const sent = (html: string, extra: Record<string, unknown> = {}): Record<string, unknown> => ({
  body: { contentType: 'html', content: `<html><body><div>text</div>${html}</body></html>` },
  sentDateTime: '2026-07-16T09:00:00Z',
  hasAttachments: false,
  ...extra,
});

// Routes each GET by path and records the order, so a test can prove a message
// was never read as readily as it can prove one was.
const scanningGraph = (byPath: Record<string, Awaited<ReturnType<GraphClient['get']>>>): { graph: GraphClient; gets: string[] } => {
  const gets: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      gets.push(path);
      return byPath[path] ?? ok({});
    },
  });
  return { graph, gets };
};

describe('get-mail-signature', () => {
  it('returns the signature from the most recent sent message that has one, and stops reading once it finds it', async () => {
    const { graph, gets } = scanningGraph({ [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }, { id: 'sent-2' }] }), [bodyOf('sent-1')]: ok(sent(SIGNATURE)) });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toEqual({
        contentType: 'text/html',
        size: 66,
        text: SIGNATURE,
        sourceMessageId: 'sent-1',
        sentDateTime: '2026-07-16T09:00:00Z',
        inlinedImages: 0,
      });
    }
    // sent-2's body is never fetched: the scan stops at the first hit, so the
    // common case costs one list call and one body read.
    expect(gets).toEqual([SCAN_PATH, bodyOf('sent-1')]);
  });

  it('keeps scanning past a sent message whose only signature is inside the quoted history', async () => {
    const { graph, gets } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }, { id: 'sent-2' }] }),
      // A reply where the author signed nothing but quoted a colleague who did.
      [bodyOf('sent-1')]: ok(sent(`${QUOTE_TAIL}${SIGNATURE}`)),
      [bodyOf('sent-2')]: ok(sent(SIGNATURE)),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ sourceMessageId: 'sent-2' });
    expect(gets.length).toBe(3);
  });

  it('reads only the message it is told to, when given one, and never touches the sent folder', async () => {
    const { graph, gets } = scanningGraph({ [bodyOf('pinned-1')]: ok(sent(SIGNATURE)) });

    const result = await execute(graph, { messageId: 'pinned-1' });

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ sourceMessageId: 'pinned-1' });
    expect(gets).toEqual([bodyOf('pinned-1')]);
  });

  it('names the Outlook-desktop limitation when no sent message carries a signature block, rather than guessing at one', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }), [bodyOf('sent-1')]: ok(sent('<div>no signature here</div>')) });

    const result = await execute(graph, {});

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      // The code is what an agent routes on; the message is for a human. The
      // 2026-08-31 audit found this rejection reaching the envelope with no
      // errorCode at all, so the sweep logged it as `ERR:err`.
      expect(result.error.code).toBe('signature_not_found');
      expect(result.error.message).toContain('Outlook desktop does not carry the marker');
      expect(result.error.message).toContain('--message-id');
    }
  });

  it('tells the caller their pinned message is the one without a signature, not that a scan came up empty', async () => {
    const { graph } = scanningGraph({ [bodyOf('pinned-1')]: ok(sent('<div>no signature here</div>')) });

    const result = await execute(graph, { messageId: 'pinned-1' });

    expect(result.ok).toBe(false);
    // Reporting "the last 1 sent message" here would be a lie: nothing was scanned.
    if (!result.ok) {
      // Same code as the scan miss: what an agent routes on is "no signature
      // here", and which of the two paths found nothing is in the message.
      expect(result.error.code).toBe('signature_not_found');
      expect(result.error.message).toBe(
        'Message pinned-1 carries no OWA signature block (`<div id="Signature">`). Mail composed in Outlook desktop does not carry the marker; pin a message sent from Outlook on the web instead.'
      );
    }
  });

  it('counts a single scanned message in the singular when it reports finding no signature', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }), [bodyOf('sent-1')]: ok(sent('<div>nothing</div>')) });

    const result = await execute(graph, {});

    if (!result.ok) expect(result.error.message).toContain('the last 1 sent message.');
  });

  it('reports an empty sent folder as the reason it found nothing', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [] }) });

    const result = await execute(graph, {});

    expect(result.ok).toBe(false);
    // The type is the machine-readable half of the contract: a caller branches on
    // it to tell "your mailbox has no signature" from "Graph is down".
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      expect(result.error.code).toBe('no_sent_messages');
      expect(result.error.message).toContain('no sent messages');
    }
  });

  it('rejects a message id that is not text, before reaching for the mailbox', async () => {
    const { graph, gets } = scanningGraph({});

    const result = await execute(graph, { messageId: 42 as unknown as string });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(gets).toEqual([]);
  });

  it('reports an unreadable sent-folder listing the same way as an empty one, rather than crashing on its shape', async () => {
    // Two shapes Graph should never send but might: no `value` key at all, and a
    // `value` that is not a list. Neither may throw on `.map`.
    for (const shape of [{}, { value: 'not-a-list' }]) {
      const { graph } = scanningGraph({ [SCAN_PATH]: ok(shape) });

      const result = await execute(graph, {});

      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.message).toContain('no sent messages');
    }
  });

  it('names an unembeddable image even when Graph gives the attachment no filename', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent('<div id="Signature"><img src="cid:x@fabrikam"></div>', { hasAttachments: true })),
      [attsOf('sent-1')]: ok({ value: [{ id: 'att-1', contentType: 'image/png', size: 3_000_000, isInline: true, contentId: 'x@fabrikam' }] }),
    });

    const result = await execute(graph, {});

    if (result.ok) expect(result.value).toMatchObject({ note: '1 inline image left as a cid: reference: image (3.0 MB)' });
  });

  it('treats an attachment list with no entries as nothing to embed', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent('<div id="Signature"><img src="cid:gone@fabrikam"></div>', { hasAttachments: true })),
      [attsOf('sent-1')]: ok({}),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ text: '<div id="Signature"><img src="cid:gone@fabrikam"></div>', inlinedImages: 0 });
  });

  it('passes a failed sent-folder read through untouched', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: err({ type: 'api_error', status: 403, message: 'Forbidden' }) });

    expect(await execute(graph, {})).toEqual(err({ type: 'api_error', status: 403, message: 'Forbidden' }));
  });

  it('embeds a logo the signature references, so the block renders on its own', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent('<div id="Signature"><img src="cid:logo@fabrikam"></div>', { hasAttachments: true })),
      [attsOf('sent-1')]: ok({ value: [{ id: 'att-1', name: 'logo.png', contentType: 'image/png', size: 120, isInline: true, contentId: 'logo@fabrikam' }] }),
      '/me/messages/sent-1/attachments/att-1': ok({ contentBytes: 'QUJD' }),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) {
      // toEqual, not toMatchObject: a successful embed reports NO note. There is
      // nothing the caller has to go and resolve themselves.
      expect(result.value).toEqual({
        contentType: 'text/html',
        size: 64,
        text: '<div id="Signature"><img src="data:image/png;base64,QUJD"></div>',
        sourceMessageId: 'sent-1',
        sentDateTime: '2026-07-16T09:00:00Z',
        inlinedImages: 1,
      });
      // The key is absent, not present-and-undefined: `note` means "something
      // needs your attention", so an always-present one would be noise.
      expect(Object.hasOwn(result.value as object, 'note')).toBe(false);
    }
  });

  it('inlines the signature logo even when Graph reports hasAttachments:false, the shape a message carries when its only images are inline', async () => {
    // Inline images (contentId + isInline) do NOT flip Graph's `hasAttachments`,
    // so a signature's logo rides in a message reported as hasAttachments:false.
    // Gating the embed on that flag left the logo a raw cid: reference, broken
    // once pasted into a fresh draft (reported 2026-07-19). The gate is the
    // block's own cid: reference, not the flag.
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent('<div id="Signature"><img src="cid:logo@fabrikam"></div>', { hasAttachments: false })),
      [attsOf('sent-1')]: ok({ value: [{ id: 'att-1', name: 'logo.png', contentType: 'image/png', size: 120, isInline: true, contentId: 'logo@fabrikam' }] }),
      '/me/messages/sent-1/attachments/att-1': ok({ contentBytes: 'QUJD' }),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ text: '<div id="Signature"><img src="data:image/png;base64,QUJD"></div>', inlinedImages: 1 });
  });

  it('names every image it could not embed, counting them, when a signature carries more than one', async () => {
    const block = '<div id="Signature"><img src="cid:a@fabrikam"><img src="cid:b@fabrikam"></div>';
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent(block, { hasAttachments: true })),
      [attsOf('sent-1')]: ok({
        value: [
          { id: 'att-1', name: 'wide.png', contentType: 'image/png', size: 3_000_000, isInline: true, contentId: 'a@fabrikam' },
          { id: 'att-2', name: 'tall.png', contentType: 'image/png', size: 4_000_000, isInline: true, contentId: 'b@fabrikam' },
        ],
      }),
    });

    const result = await execute(graph, {});

    if (result.ok) expect(result.value).toMatchObject({ inlinedImages: 0, note: '2 inline images left as a cid: reference: wide.png (3.0 MB), tall.png (4.0 MB)' });
  });

  it('omits the sent date rather than reporting an empty one when Graph does not give it', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok({ body: { contentType: 'html', content: `<html><body>${SIGNATURE}</body></html>` } }),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) expect(Object.hasOwn(result.value as object, 'sentDateTime')).toBe(false);
  });

  it('passes a failed message read through untouched, rather than scanning past a message it could not see', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }, { id: 'sent-2' }] }),
      [bodyOf('sent-1')]: err({ type: 'api_error', status: 503, message: 'ServiceUnavailable' }),
    });

    // Scanning on would report "no signature found", which is a different and
    // wrong answer: the mailbox may well have one in the message that failed.
    expect(await execute(graph, {})).toEqual(err({ type: 'api_error', status: 503, message: 'ServiceUnavailable' }));
  });

  it('scans past a sent message that comes back with no body at all', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }, { id: 'sent-2' }] }),
      [bodyOf('sent-1')]: ok({ sentDateTime: '2026-07-16T09:00:00Z' }),
      [bodyOf('sent-2')]: ok(sent(SIGNATURE)),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ sourceMessageId: 'sent-2' });
  });

  it('counts the messages it actually scanned when it reports finding no signature', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }, { id: 'sent-2' }] }),
      [bodyOf('sent-1')]: ok(sent('<div>nothing</div>')),
      [bodyOf('sent-2')]: ok(sent('<div>nothing</div>')),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('the last 2 sent messages');
  });

  it('does not fetch bytes for an inline image the signature never references', async () => {
    const { graph, gets } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent(SIGNATURE, { hasAttachments: true })),
      [attsOf('sent-1')]: ok({ value: [{ id: 'att-1', name: 'chart.png', contentType: 'image/png', size: 120, isInline: true, contentId: 'chart@fabrikam' }] }),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ inlinedImages: 0 });
    expect(gets.some((p) => p.includes('/attachments/att-1'))).toBe(false);
  });

  it('keeps the cid reference and names the image when it is too large to embed, rather than destroying the reference', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent('<div id="Signature"><img src="cid:big@fabrikam"></div>', { hasAttachments: true })),
      [attsOf('sent-1')]: ok({ value: [{ id: 'att-1', name: 'banner.png', contentType: 'image/png', size: 3_000_000, isInline: true, contentId: 'big@fabrikam' }] }),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) {
      // The cid: ref survives. A placeholder would be a lie inside a signature
      // destined for a draft: the caller can still fetch the bytes themselves.
      expect(result.value).toMatchObject({
        text: '<div id="Signature"><img src="cid:big@fabrikam"></div>',
        inlinedImages: 0,
        note: '1 inline image left as a cid: reference: banner.png (3.0 MB)',
      });
    }
  });

  it('still returns the signature when the attachment list cannot be read, saying what was lost', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent('<div id="Signature"><img src="cid:logo@fabrikam"></div>', { hasAttachments: true })),
      [attsOf('sent-1')]: err({
        type: 'api_error',
        status: 500,
        message: 'boom',
      }),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ inlinedImages: 0, note: 'inline images could not be listed, so any cid: references are unresolved' });
  });

  it('still returns the signature when the attachment list comes back in a shape it cannot read, saying what was lost', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      [bodyOf('sent-1')]: ok(sent('<div id="Signature"><img src="cid:logo@fabrikam"></div>', { hasAttachments: true })),
      [attsOf('sent-1')]: ok({ value: 'not-an-array' }),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toMatchObject({ inlinedImages: 0, note: 'the inline-image list came back in an unreadable shape, so any cid: references are unresolved' });
    }
  });

  it('skips the attachments call entirely for a signature on a message that has none', async () => {
    const { graph, gets } = scanningGraph({ [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }), [bodyOf('sent-1')]: ok(sent(SIGNATURE)) });

    await execute(graph, {});

    expect(gets.some((p) => p.includes('/attachments'))).toBe(false);
  });

  it('measures the signature in utf-8 bytes, not in characters', async () => {
    const block = '<div id="Signature">Café €</div>';
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }), [bodyOf('sent-1')]: ok(sent(block)) });

    const result = await execute(graph, {});

    // 31 characters, 35 bytes: é is 2 and € is 3.
    if (result.ok) expect(result.value).toMatchObject({ size: 35 });
  });
});
