import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './get-mail-signature.ts';

const SCAN_PATH = '/me/mailFolders/sentitems/messages?$top=10&$orderby=sentDateTime%20desc&$select=id,sentDateTime';
const SIGNATURE = '<div id="Signature"><div>Robin Chen</div><div>Fabrikam</div></div>';
const QUOTE_TAIL = '<div id="appendonsend"></div><div id="divRplyFwdMsg"><b>From:</b> Alex Kim<br><b>Sent:</b> Monday</div>';

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
    const { graph, gets } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }, { id: 'sent-2' }] }),
      '/me/messages/sent-1?$select=body,sentDateTime': ok(sent(SIGNATURE)),
    });

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
    expect(gets).toEqual([SCAN_PATH, '/me/messages/sent-1?$select=body,sentDateTime']);
  });

  it('keeps scanning past a sent message whose only signature is inside the quoted history', async () => {
    const { graph, gets } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }, { id: 'sent-2' }] }),
      // A reply where the author signed nothing but quoted a colleague who did.
      '/me/messages/sent-1?$select=body,sentDateTime': ok(sent(`${QUOTE_TAIL}${SIGNATURE}`)),
      '/me/messages/sent-2?$select=body,sentDateTime': ok(sent(SIGNATURE)),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ sourceMessageId: 'sent-2' });
    expect(gets.length).toBe(3);
  });

  it('reads only the message it is told to, when given one, and never touches the sent folder', async () => {
    const { graph, gets } = scanningGraph({
      '/me/messages/pinned-1?$select=body,sentDateTime': ok(sent(SIGNATURE)),
    });

    const result = await execute(graph, { messageId: 'pinned-1' });

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toMatchObject({ sourceMessageId: 'pinned-1' });
    expect(gets).toEqual(['/me/messages/pinned-1?$select=body,sentDateTime']);
  });

  it('names the Outlook-desktop limitation when no sent message carries a signature block, rather than guessing at one', async () => {
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      '/me/messages/sent-1?$select=body,sentDateTime': ok(sent('<div>no signature here</div>')),
    });

    const result = await execute(graph, {});

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.message).toContain('Outlook desktop does not carry the marker');
      expect(result.error.message).toContain('--message-id');
    }
  });

  it('reports an empty sent folder as the reason it found nothing', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [] }) });

    const result = await execute(graph, {});

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('no sent messages');
  });

  it('passes a failed sent-folder read through untouched', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: err({ type: 'api_error', status: 403, message: 'Forbidden' }) });

    expect(await execute(graph, {})).toEqual(err({ type: 'api_error', status: 403, message: 'Forbidden' }));
  });

  it('measures the signature in utf-8 bytes, not in characters', async () => {
    const block = '<div id="Signature">Café €</div>';
    const { graph } = scanningGraph({
      [SCAN_PATH]: ok({ value: [{ id: 'sent-1' }] }),
      '/me/messages/sent-1?$select=body,sentDateTime': ok(sent(block)),
    });

    const result = await execute(graph, {});

    // 31 characters, 35 bytes: é is 2 and € is 3.
    if (result.ok) expect(result.value).toMatchObject({ size: 35 });
  });
});
