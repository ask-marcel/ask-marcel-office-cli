// The write cases of the command suite: the PDF converters, which upload a temporary
// OneDrive file. Moved out of commands.test.ts with their text unchanged (package split,
// step 16), so the write package takes this file whole; the helpers are copies of that file's.
import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../../domain/access-token.ts';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { AuthManager } from '../../test-helpers/auth-manager-fake.ts';
import type { FetchFn, GraphError } from '../../infra/graph-client.ts';
import { createGraphClient } from '../../infra/graph-client.ts';
import { fakeAuthManager } from '../../test-helpers/auth-manager-fake.ts';
import type { CommandExecute } from './command-types.ts';
import * as convertMailAttachmentToPdf from './convert-mail-attachment-to-pdf.ts';
import * as convertCalendarEventAttachmentToPdf from './convert-calendar-event-attachment-to-pdf.ts';

const cmdMap: Record<string, { execute: CommandExecute }> = {
  'convert-mail-attachment-to-pdf': convertMailAttachmentToPdf,
  'convert-calendar-event-attachment-to-pdf': convertCalendarEventAttachmentToPdf,
};

const fakeAuth = (): AuthManager =>
  fakeAuthManager({
    getAccessToken: async () => ok(accessTokenUnsafe('test-token')),
    getElevatedAccessToken: async () => ok(accessTokenUnsafe('test-elevated-token')),
    getGuestAccessToken: async () => ok(accessTokenUnsafe('test-guest-token')),
    getChatsvcaggAccessToken: async () => ok(accessTokenUnsafe('test-chatsvcagg-token')),
    getIc3AccessToken: async () => ok(accessTokenUnsafe('test-ic3-token')),
  });

type FakeFetch = ((url: string, init?: RequestInit) => Promise<Response>) & { lastUrl: string | null; lastBody: string | null };

const fakeFetch = (body: unknown): FakeFetch => {
  let lastUrl: string | null = null;
  let lastBody: string | null = null;
  const fn = async (url: string, init?: RequestInit): Promise<Response> => {
    lastUrl = url;
    lastBody = typeof init?.body === 'string' ? init.body : null;
    return new Response(JSON.stringify(body), { headers: { 'content-type': 'application/json' } });
  };
  Object.defineProperty(fn, 'lastUrl', { get: () => lastUrl });
  Object.defineProperty(fn, 'lastBody', { get: () => lastBody });
  return fn as FakeFetch;
};

const callCommand = async (name: string, params: Record<string, string>, responseBody: unknown): Promise<Result<unknown, GraphError>> => {
  const cmd = cmdMap[name];
  if (!cmd) throw new Error(`command not found: ${name}`);
  const fetchFn = fakeFetch(responseBody);
  const graph = createGraphClient(fakeAuth(), fetchFn);
  return cmd.execute(graph, params);
};

const capturedUrl = async (name: string, params: Record<string, string>): Promise<string> => {
  const cmd = cmdMap[name];
  if (!cmd) throw new Error(`command not found: ${name}`);
  const fetchFn = fakeFetch({ ok: true });
  const graph = createGraphClient(fakeAuth(), fetchFn);
  await cmd.execute(graph, params);
  return fetchFn.lastUrl ?? '';
};

type CommandFixture = { readonly name: string; readonly params: Record<string, string>; readonly responseBody?: unknown };

describe('commands', () => {
  it('convert-mail-attachment-to-pdf uploads a fileAttachment, converts via ?format=pdf, follows the CDN redirect internally to inline the PDF bytes, then deletes the temp item', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/a1')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: btoa('docx-bytes') });
      }
      if (url.includes(':/content') && init?.method === 'PUT') {
        return Response.json({ id: 'temp-i1', name: 'plan-temp' });
      }
      if (url.endsWith('/content?format=pdf')) {
        return Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/plan.pdf' });
      }
      if (url === 'https://contoso.sharepoint.com/cdn/plan.pdf') {
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      if (url.endsWith('/items/temp-i1') && init?.method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'a1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
    expect(calls.some((c) => c.method === 'DELETE')).toBe(true);
    expect(calls.some((c) => c.url.endsWith('/content?format=pdf'))).toBe(true);
  });

  it('convert-mail-attachment-to-pdf turns a 406 from the format=pdf transform (an input it refuses — xlsx observed live) into an actionable error pointing at the markdown fallback, NOT the misleading "endpoint may have moved" text, and still cleans up the temp item', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/a1')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'budget.xlsx', contentBytes: btoa('xlsx-bytes') });
      if (url.includes(':/content') && init?.method === 'PUT') return Response.json({ id: 'temp-i1' });
      if (url.endsWith('/content?format=pdf')) return new Response(null, { status: 406 });
      if (url.endsWith('/items/temp-i1') && init?.method === 'DELETE') return new Response(null, { status: 204 });
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'a1' });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.message).toContain('convert-mail-attachment-to-markdown');
      expect(result.error.message).toContain('get-mail-attachment');
      expect(result.error.message).toMatch(/406|refused|could not/i);
      expect(result.error.message).not.toContain('endpoint may have moved');
    }
    expect(calls.some((c) => c.method === 'DELETE')).toBe(true); // temp item cleaned up even on failure
  });

  it('convert-mail-attachment-to-pdf deletes the empty `.ask-marcel-temp` parent folder after a successful upload-convert-delete cycle (folder used to linger at OneDrive root)', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/aPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: btoa('docx') });
      }
      if (url.includes(':/content') && init?.method === 'PUT') {
        return Response.json({ id: 'temp-i7', name: 'plan-temp' });
      }
      if (url.endsWith('/content?format=pdf')) {
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      if (url.endsWith('/items/temp-i7') && init?.method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      // Cleanup: children GET returns empty → folder delete fires.
      if (url.includes('/.ask-marcel-temp:/children')) {
        return Response.json({ value: [] });
      }
      if (url.endsWith('/.ask-marcel-temp') && init?.method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm7', attachmentId: 'aPdf' });
    expect(result.ok).toBe(true);
    // Both the file delete AND the folder delete happened.
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/items/temp-i7'))).toBe(true);
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/.ask-marcel-temp'))).toBe(true);
    // The cleanup probed children first.
    expect(calls.some((c) => c.url.includes('/.ask-marcel-temp:/children'))).toBe(true);
  });

  it('convert-mail-attachment-to-pdf also sweeps the legacy un-dotted `ask-marcel-temp` folder when empty (pre-rename orphan found at a real tenant root)', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/aPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: btoa('docx') });
      }
      if (url.includes(':/content') && init?.method === 'PUT') return Response.json({ id: 'temp-i8', name: 'plan-temp' });
      if (url.endsWith('/content?format=pdf')) return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      if (url.endsWith('/items/temp-i8') && init?.method === 'DELETE') return new Response(null, { status: 204 });
      // BOTH temp folders report empty → both get the best-effort delete.
      if (url.includes('-temp:/children')) return Response.json({ value: [] });
      if (url.endsWith('-temp') && init?.method === 'DELETE') return new Response(null, { status: 204 });
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm8', attachmentId: 'aPdf' });
    expect(result.ok).toBe(true);
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/.ask-marcel-temp'))).toBe(true);
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/root:/ask-marcel-temp'))).toBe(true); // the legacy un-dotted orphan
  });

  it('convert-mail-attachment-to-pdf does NOT delete the parent folder when it still contains other files (concurrent invocation safety)', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/aPdf2')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'b.docx', contentBytes: btoa('docx') });
      }
      if (url.includes(':/content') && init?.method === 'PUT') {
        return Response.json({ id: 'temp-i8', name: 'b-temp' });
      }
      if (url.endsWith('/content?format=pdf')) {
        return new Response(new Uint8Array([0x25, 0x50]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      if (url.endsWith('/items/temp-i8') && init?.method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      // Cleanup: children GET returns one other file → folder delete must NOT fire.
      if (url.includes('/.ask-marcel-temp:/children')) {
        return Response.json({ value: [{ id: 'other-concurrent-upload' }] });
      }
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm8', attachmentId: 'aPdf2' });
    expect(result.ok).toBe(true);
    // File delete happened, but the folder delete did NOT — concurrent file survives.
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/items/temp-i8'))).toBe(true);
    expect(calls.every((c) => !(c.method === 'DELETE' && c.url.endsWith('/.ask-marcel-temp')))).toBe(true);
  });

  it('convert-mail-attachment-to-pdf short-circuits to a raw-bytes envelope for plain-text source extensions', async () => {
    const fetchFn: FetchFn = async (url, init) => {
      if (url.endsWith('/attachments/aText')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'README.md', contentBytes: btoa('# Hello') });
      }
      throw new Error(`unexpected fetch ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; size: number; base64: string; note: string; passthrough: boolean };
      expect(v.contentType).toBe('text/plain');
      expect(v.note).toContain('pre-checked source');
      expect(v.passthrough).toBe(true); // so --output-path guard blocks writing text into a .pdf
      expect(atob(v.base64)).toBe('# Hello');
    }
  });

  it('convert-mail-attachment-to-pdf short-circuits a pdf fileAttachment to its raw bytes (no upload-convert dance, no Graph format=pdf call)', async () => {
    const calls: Array<{ method: string; url: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ method: init?.method ?? 'GET', url });
      if (url.endsWith('/attachments/aPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf', contentBytes: btoa('%PDF-fake') });
      }
      throw new Error(`unexpected fetch ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPdf' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string; note: string };
      expect(v.contentType).toBe('application/pdf');
      expect(v.note).toContain('pre-checked source');
      expect(atob(v.base64)).toBe('%PDF-fake');
    }
    expect(calls.some((c) => c.method === 'PUT')).toBe(false);
    expect(calls.some((c) => c.url.includes('format=pdf'))).toBe(false);
  });
});

describe('calendar event attachments', () => {
  it('convert-calendar-event-attachment-to-pdf short-circuits a pdf event attachment to raw bytes through the shared pipeline', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.includes('/me/events/e1/attachments/a1')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pdf', contentBytes: btoa('PDFB') });
      throw new Error(`unexpected ${url}`);
    };
    const result = await cmdMap['convert-calendar-event-attachment-to-pdf']?.execute(createGraphClient(fakeAuth(), fetchFn), { eventId: 'e1', attachmentId: 'a1' });
    expect(result?.ok).toBe(true);
    if (result?.ok) {
      const v = result.value as { contentType: string; note?: string };
      expect(v.contentType).toBe('application/pdf');
      expect(String(v.note)).toContain('raw bytes');
    }
  });

  it('convert-calendar-event-attachment-to-pdf returns a validation_error when eventId is missing', async () => {
    const result = await cmdMap['convert-calendar-event-attachment-to-pdf']?.execute(createGraphClient(fakeAuth(), fakeFetch({})), { attachmentId: 'a1' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.type).toBe('validation_error');
  });
});

const allCommandFixtures: CommandFixture[] = [
  {
    name: 'convert-calendar-event-attachment-to-pdf',
    params: { eventId: 'e1', attachmentId: 'a1' },
    responseBody: { '@odata.type': '#microsoft.graph.fileAttachment', name: 'note.txt', contentBytes: 'aGk=' },
  },
];

describe('all commands schema acceptance', () => {
  it.each(allCommandFixtures)('$name accepts valid params', async ({ name, params, responseBody }) => {
    const result = await callCommand(name, params, responseBody ?? { ok: true });
    expect(result.ok).toBe(true);
  });
});

describe('command schema rejection', () => {
  const rejectCases: Array<{ name: string; params: Record<string, string> }> = [{ name: 'convert-mail-attachment-to-pdf', params: { messageId: 'm1' } }];

  it.each(rejectCases)('$name rejects missing required params as a validation_error Result', async ({ name, params }) => {
    const cmd = cmdMap[name];
    const fetchFn = fakeFetch({ ok: true });
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, params);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});

const pathFixtures: Array<{ name: string; params: Record<string, string>; expectedPath: string }> = [
  { name: 'convert-calendar-event-attachment-to-pdf', params: { eventId: 'e1', attachmentId: 'a1' }, expectedPath: '/me/events/e1/attachments/a1' },
];

describe('all commands build correct Graph URL', () => {
  it.each(pathFixtures)('$name calls $expectedPath', async ({ name, params, expectedPath }) => {
    const url = await capturedUrl(name, params);
    expect(url).toBe(`https://graph.microsoft.com/v1.0${expectedPath}`);
  });
});
