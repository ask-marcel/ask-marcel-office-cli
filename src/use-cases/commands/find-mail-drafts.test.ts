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

  // A draft matches when it shares AT LEAST ONE recipient with the wanted list, not
  // only when it shares EVERY one. A reply you are about to send to kim + alex must
  // still surface a draft already addressed to kim alone, so it is revised not duplicated.
  it('matches a draft sharing one of several wanted recipients, not only a draft sharing all of them', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [draftItem('a', 'RE: Contoso', ['kim@example.com'], 'conv-x')] }) });

    const result = await execute(graph, { subject: 'Contoso', toRecipients: 'kim@example.com,alex@example.com' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: Array<{ id: string }> };
    expect(value.matches.map((m) => m.id)).toEqual(['a']);
  });

  // Subject normalization collapses internal whitespace runs to a single space AND
  // trims the ends before comparing, so cosmetic spacing never hides a duplicate:
  // 'Contoso  Q3' (double space) and 'Contoso Q3 ' (trailing space) both match 'Contoso Q3'.
  it('matches drafts whose subject differs from the wanted one only in whitespace', async () => {
    const drafts = [draftItem('a', 'Contoso  Q3', ['kim@example.com'], 'conv-x'), draftItem('b', 'Contoso Q3 ', ['kim@example.com'], 'conv-y')];
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: drafts }) });

    const result = await execute(graph, { subject: 'Contoso Q3' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: Array<{ id: string }> };
    expect(value.matches.map((m) => m.id)).toEqual(['a', 'b']);
  });

  // A draft recipient can arrive with no emailAddress object at all, or with one that
  // carries no address. Neither may crash address extraction; the draft still matches
  // on a recipient that IS well-formed.
  it('tolerates recipients missing their emailAddress or address and still matches on a well-formed one', async () => {
    const draft = { id: 'a', subject: 'RE: Contoso', conversationId: 'conv-x', toRecipients: [{}, { emailAddress: {} }, { emailAddress: { address: 'kim@example.com' } }] };
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [draft] }) });

    const result = await execute(graph, { subject: 'Contoso', toRecipients: 'kim@example.com' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: Array<{ id: string }> };
    expect(value.matches.map((m) => m.id)).toEqual(['a']);
  });

  // A pasted `--to-recipients " kim@example.com "` carries surrounding whitespace; it is
  // trimmed before matching so it still shares an address with a draft addressed to kim.
  it('trims surrounding whitespace on a wanted recipient before matching', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [draftItem('a', 'RE: Contoso', ['kim@example.com'], 'conv-x')] }) });

    const result = await execute(graph, { subject: 'Contoso', toRecipients: ' kim@example.com ' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: Array<{ id: string }> };
    expect(value.matches.map((m) => m.id)).toEqual(['a']);
  });

  // A Graph read failure must reach the caller with ITS OWN code, not be masked as a
  // generic "unreadable shape" error, so the caller can tell a transient 503 from bad data.
  it('passes the specific Graph error code straight back to the caller', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: err({ type: 'api_error', status: 503, code: 'service_unavailable', message: 'Graph is down' }) });

    const result = await execute(graph, { subject: 'Contoso' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    if (result.error.type !== 'api_error') return;
    expect(result.error.code).toBe('service_unavailable');
  });

  // An unreadable Drafts listing carries the machine-readable `drafts_list_unreadable`
  // code so an agent routes on the code rather than substring-matching the message.
  it('stamps drafts_list_unreadable as the code when the listing shape is not an object', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: 'not-an-array' }) });

    const result = await execute(graph, { subject: 'Contoso' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    if (result.error.type !== 'api_error') return;
    expect(result.error.code).toBe('drafts_list_unreadable');
  });

  // An empty Drafts folder (a listing with no `value` array) reports scanned 0 and no
  // matches, never a phantom scanned count.
  it('reports scanned 0 and no matches for an empty drafts listing', async () => {
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({}) });

    const result = await execute(graph, { subject: 'Contoso' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: unknown[]; scanned: number };
    expect(value.matches).toEqual([]);
    expect(value.scanned).toBe(0);
  });

  // A matched draft with no conversationId must not inject `undefined` into the deduped
  // conversationIds union — an agent iterating it would otherwise hit an undefined id.
  it('excludes a matched draft that has no conversationId from the conversationIds union', async () => {
    const draft = { id: 'a', subject: 'RE: Contoso', toRecipients: [{ emailAddress: { address: 'kim@example.com' } }] };
    const { graph } = scanningGraph({ [SCAN_PATH]: ok({ value: [draft] }) });

    const result = await execute(graph, { subject: 'Contoso' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { matches: Array<{ id: string }>; conversationIds: string[] };
    expect(value.matches.map((m) => m.id)).toEqual(['a']);
    // toHaveLength(0), not toEqual([]): Bun's toEqual treats [undefined] as equal to [], which
    // would let an `undefined` conversationId slip through this very assertion.
    expect(value.conversationIds).toHaveLength(0);
  });
});
