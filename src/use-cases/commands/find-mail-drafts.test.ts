import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './find-mail-drafts.ts';

const SCAN_PATH =
  '/me/mailFolders/drafts/messages?$top=50&$orderby=lastModifiedDateTime%20desc&$select=id,subject,toRecipients,ccRecipients,conversationId,lastModifiedDateTime,webLink';

const draftItem = (id: string, subject: string, addresses: ReadonlyArray<string>, conversationId: string): Record<string, unknown> => ({
  id,
  subject,
  conversationId,
  toRecipients: addresses.map((address) => ({ emailAddress: { address } })),
});

// Routes each GET by path and records the order, so a test can prove the exact
// scan URL was issued.
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

describe('find-mail-drafts', () => {
  it('scans the drafts folder newest-first and finds thread drafts a conversationId filter would miss', async () => {
    const drafts = [
      draftItem('a', 'RE: Contoso Q3', ['kim@example.com'], 'conv-x'),
      draftItem('b', 'FW: Contoso Q3', ['kim@example.com'], 'conv-y'),
      draftItem('c', 'RE: Fabrikam', ['kim@example.com'], 'conv-z'),
    ];
    const { graph, gets } = scanningGraph({ [SCAN_PATH]: ok({ value: drafts }) });

    const result = await execute(graph, { subject: 'Contoso Q3' });

    expect(gets[0]).toBe(SCAN_PATH);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: Array<{ id: string }>; conversationIds: string[]; scanned: number; scanLimit: number };
    expect(value.matches.map((m) => m.id)).toEqual(['a', 'b']);
    expect(value.conversationIds).toEqual(['conv-x', 'conv-y']);
    expect(value.scanned).toBe(3);
    expect(value.scanLimit).toBe(50);
  });

  it('narrows matches to a shared recipient when --to-recipients is given', async () => {
    const drafts = [draftItem('a', 'RE: Contoso', ['kim@example.com'], 'conv-x'), draftItem('b', 'RE: Contoso', ['other@example.com'], 'conv-y')];
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: drafts }) });

    const result = await execute(graph, { subject: 'Contoso', toRecipients: 'kim@example.com' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: Array<{ id: string }> };
    expect(value.matches.map((m) => m.id)).toEqual(['a']);
  });

  it('reports zero matches with the scan count when nothing is on the thread', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [draftItem('a', 'RE: Fabrikam', ['kim@example.com'], 'conv-z')] }) });

    const result = await execute(graph, { subject: 'Contoso' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: unknown[]; scanned: number };
    expect(value.matches).toEqual([]);
    expect(value.scanned).toBe(1);
  });

  it('passes a Graph read error straight back to the caller', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: err({ type: 'api_error', status: 503, code: 'service_unavailable', message: 'Graph is down' }) });

    const result = await execute(graph, { subject: 'Contoso' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
  });

  it('surfaces an api_error when the drafts list comes back in an unreadable shape', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: 'not-an-array' }) });

    const result = await execute(graph, { subject: 'Contoso' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
  });

  it('rejects a call with no subject', async () => {
    const { graph } = scanningGraph({});

    const result = await execute(graph, {});

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });
});
