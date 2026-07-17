import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './get-mail-signature.ts';

const SCAN_PATH = '/me/mailFolders/sentitems/messages?$top=10&$orderby=sentDateTime%20desc&$select=id,sentDateTime';
const SIGNATURE = '<div id="Signature"><div>Robin Chen</div><div>Fabrikam</div></div>';
const QUOTE_TAIL = '<div id="appendonsend"></div><div id="divRplyFwdMsg"><b>From:</b> Alex Kim<br><b>Sent:</b> Monday</div>';
const bodyOf = (id: string): string => `/me/messages/${id}?$select=body,sentDateTime`;

const sent = (html: string): Record<string, unknown> => ({
  body: { contentType: 'html', content: `<html><body><div>text</div>${html}</body></html>` },
  sentDateTime: '2026-07-16T09:00:00Z',
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
      expect(result.error.message).toContain('Outlook desktop does not carry the marker');
      expect(result.error.message).toContain('--message-id');
    }
  });

  it('tells the caller their pinned message is the one without a signature, not that a scan came up empty', async () => {
    const { graph } = scanningGraph({ [bodyOf('pinned-1')]: ok(sent('<div>no signature here</div>')) });

    const result = await execute(graph, { messageId: 'pinned-1' });

    expect(result.ok).toBe(false);
    // Reporting "the last 1 sent message" here would be a lie: nothing was scanned.
    if (!result.ok)
      expect(result.error.message).toBe(
        'Message pinned-1 carries no OWA signature block (`<div id="Signature">`). Mail composed in Outlook desktop does not carry the marker; pin a message sent from Outlook on the web instead.'
      );
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

  it('passes a failed sent-folder read through untouched', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: err({ type: 'api_error', status: 403, message: 'Forbidden' }) });

    expect(await execute(graph, {})).toEqual(err({ type: 'api_error', status: 403, message: 'Forbidden' }));
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

  it('measures the signature in utf-8 bytes, not in characters', async () => {
    const block = '<div id="Signature">Café €</div>';
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }), [bodyOf('sent-1')]: ok(sent(block)) });

    const result = await execute(graph, {});

    // 31 characters, 35 bytes: é is 2 and € is 3.
    if (result.ok) expect(result.value).toMatchObject({ size: 35 });
  });
});
