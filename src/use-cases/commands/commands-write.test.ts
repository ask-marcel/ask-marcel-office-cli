// The write cases of the command suite: the PDF converters, which upload a temporary
// OneDrive file. Moved out of commands.test.ts with their text unchanged (package split,
// step 16), so the write package takes this file whole; the helpers are copies of that file's.
import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { TeamsRegion } from '@ask-marcel/office-core/domain/teams-region.ts';
import type { TokenSource } from '@ask-marcel/office-core/use-cases/ports/token-source.ts';
import type { FetchFn, GraphError } from '../../infra/graph-client.ts';
import { createTokenSourceGraphClient as createGraphClient } from '../../infra/graph-client.ts';
import type { CommandExecute } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import * as convertMailAttachmentToPdf from './convert-mail-attachment-to-pdf.ts';
import * as convertCalendarEventAttachmentToPdf from './convert-calendar-event-attachment-to-pdf.ts';

const cmdMap: Record<string, { execute: CommandExecute }> = {
  'convert-mail-attachment-to-pdf': convertMailAttachmentToPdf,
  'convert-calendar-event-attachment-to-pdf': convertCalendarEventAttachmentToPdf,
};

const fakeAuth = (): TokenSource => ({
  graphToken: async (tier) => ok(accessTokenUnsafe(tier === 'elevated' ? 'test-elevated-token' : 'test-token')),
  guestToken: async () => ok(accessTokenUnsafe('test-guest-token')),
  substrateToken: async (tier) => ok(accessTokenUnsafe(`test-${tier}-token`)),
  substrateRegion: async () => ok('emea' as TeamsRegion),
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

  it('convert-mail-attachment-to-pdf short-circuits a pdf referenceAttachment to its raw bytes via /content (no format=pdf), follows the CDN redirect internally', async () => {
    let formatPdfCalled = false;
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/report.pdf' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-pdf', name: 'report.pdf', parentReference: { driveId: 'd1' } });
      }
      if (url.endsWith('/drives/d1/items/i-pdf/content')) {
        return Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/report.pdf' });
      }
      if (url === 'https://contoso.sharepoint.com/cdn/report.pdf') {
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      if (url.includes('format=pdf')) {
        formatPdfCalled = true;
        return new Response(null, { status: 500 });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefPdf' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
    expect(formatPdfCalled).toBe(false);
  });

  it('convert-mail-attachment-to-pdf tags a plain-text referenceAttachment short-circuit as passthrough so --output-path cannot write it into a .pdf', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefTxt')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/notes.txt' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-txt', name: 'notes.txt', parentReference: { driveId: 'd1' } });
      }
      if (url.endsWith('/drives/d1/items/i-txt/content')) {
        return new Response('plain text body', { status: 200, headers: { 'content-type': 'text/plain' } });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aRefTxt' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; passthrough: boolean };
      expect(v.contentType).toBe('text/plain');
      expect(v.passthrough).toBe(true);
    }
  });

  it('convert-mail-attachment-to-pdf resolves a referenceAttachment via /shares/{token}/driveItem, converts in place, and inlines the PDF bytes (CDN redirect followed internally)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRef')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/q3.docx' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-q3', name: 'q3.docx', parentReference: { driveId: 'd1' } });
      }
      if (url.endsWith('/drives/d1/items/i-q3/content?format=pdf')) {
        return Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/q3.pdf' });
      }
      if (url === 'https://contoso.sharepoint.com/cdn/q3.pdf') {
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRef' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('convert-mail-attachment-to-pdf rejects an image fileAttachment with a friendly hint pointing at get-mail-attachment + a vision-capable model', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aPng')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'screenshot.png', contentBytes: btoa('fake-png-bytes') });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPng' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('png attachment is an image');
      expect(result.error.message).toContain('get-mail-attachment');
      expect(result.error.message).toContain('vision-capable model');
    }
  });

  it('convert-mail-attachment-to-pdf rejects an image referenceAttachment with the same friendly hint (no upload-then-fail dance)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefImg')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/diagram.svg' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-svg', name: 'diagram.svg', parentReference: { driveId: 'd1' } });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefImg' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('svg attachment is an image');
    }
  });

  it('convert-mail-attachment-to-pdf rejects itemAttachment with a clear unsupported error pointing at the markdown variant', async () => {
    const fetchFn: FetchFn = async () =>
      Response.json({ '@odata.type': '#microsoft.graph.itemAttachment', item: { '@odata.type': '#microsoft.graph.message', subject: 'embedded' } });
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aItem' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(400);
      expect(result.error.message).toContain('convert-mail-attachment-to-markdown');
    }
  });

  it('convert-mail-attachment-to-pdf returns api_error when @odata.type is missing (FIX #3 discriminator guard)', async () => {
    const fetchFn: FetchFn = async () => Response.json({ name: 'no-type.docx' });
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aBad' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('missing @odata.type discriminator');
    }
  });

  it('convert-mail-attachment-to-pdf returns api_error when the @odata.type is unknown', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.weirdNewType' });
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aWeird' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('unsupported attachment type');
    }
  });

  // --- mutation hardening: unconditional error assertions (the `if (error.type === ...)`
  // guards above let type/object/message mutants survive — LESSONS.md), the full
  // IMAGE_EXTENSIONS set, and safeExtension's temp-filename extension via the PUT path. ---
  const pdfCmd = (): { execute: CommandExecute } => {
    const c = cmdMap['convert-mail-attachment-to-pdf'];
    if (!c) throw new Error('convert-mail-attachment-to-pdf not registered');
    return c;
  };
  const fileAtt =
    (name: string): FetchFn =>
    async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: btoa('bytes') });
      throw new Error(`unexpected ${url}`);
    };
  // Records every PUT (url + decoded body + content-type) and every DELETE url,
  // and serves the convert + cleanup-children responses. `children` controls the
  // cleanup probe body so the `(value ?? [])` default is observable; `putStatus`
  // forces an upload failure.
  type PutRecord = { readonly url: string; readonly body: string; readonly contentType: string };
  const uploadRecorder =
    (name: string, opts: { puts?: PutRecord[]; deletes?: string[]; children?: unknown; putStatus?: number }): FetchFn =>
    async (url, init) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: btoa('PDFDATA') });
      if (init?.method === 'PUT') {
        const decoded = init.body instanceof Uint8Array ? new TextDecoder().decode(init.body) : String(init.body);
        const ct = (init.headers as Record<string, string> | undefined)?.['content-type'] ?? '';
        opts.puts?.push({ url, body: decoded, contentType: ct });
        if (opts.putStatus !== undefined && opts.putStatus >= 400) return new Response('nope', { status: opts.putStatus });
        return Response.json({ id: 'temp-h1' });
      }
      if (url.endsWith('/content?format=pdf')) return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      if (url.includes('/.ask-marcel-temp:/children')) return Response.json(opts.children ?? { value: [] });
      if (init?.method === 'DELETE') {
        opts.deletes?.push(url);
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected ${init?.method ?? 'GET'} ${url}`);
    };

  it('convert-mail-attachment-to-pdf rejects every IMAGE_EXTENSIONS member with a 415 image hint', async () => {
    for (const ext of ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'svg', 'ico', 'heic', 'heif']) {
      const result = await pdfCmd().execute(createGraphClient(fakeAuth(), fileAtt(`pic.${ext}`)), { messageId: 'm1', attachmentId: 'a1' });
      expect(result.ok).toBe(false);
      if (result.ok) continue;
      expect(result.error.type).toBe('api_error');
      expect(result.error.message).toContain(`${ext} attachment is an image`);
      if (result.error.type === 'api_error') expect(result.error.status).toBe(415);
    }
  });

  it('convert-mail-attachment-to-pdf uploads the decoded bytes under a UUID temp name (safe ext, octet-stream content-type)', async () => {
    const docxPuts: PutRecord[] = [];
    await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('report.docx', { puts: docxPuts })), { messageId: 'm1', attachmentId: 'a1' });
    expect(docxPuts[0]?.url).toContain('.ask-marcel-temp/');
    expect(docxPuts[0]?.url.endsWith('.docx') || (docxPuts[0]?.url.includes('.docx:') ?? false)).toBe(true);
    expect(docxPuts[0]?.url).not.toContain('report.docx'); // never the attacker filename
    expect(docxPuts[0]?.body).toBe('PDFDATA'); // decodeBase64 round-trips the payload exactly (kills the loop mutants)
    expect(docxPuts[0]?.contentType).toBe('application/octet-stream'); // kills the L72 content-type literal

    const weirdPuts: PutRecord[] = [];
    await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('archive.spreadsheet', { puts: weirdPuts })), { messageId: 'm1', attachmentId: 'a1' });
    expect(weirdPuts[0]?.url.endsWith('.bin') || (weirdPuts[0]?.url.includes('.bin:') ?? false)).toBe(true); // ext > 8 chars → bin

    const noExtPuts: PutRecord[] = [];
    await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('README', { puts: noExtPuts })), { messageId: 'm1', attachmentId: 'a1' });
    // 'README' is ≤8 alnum chars: a broken no-dot guard would slice the whole
    // name → '.readme'. The correct guard returns 'bin'.
    expect(noExtPuts[0]?.url.endsWith('.bin') || (noExtPuts[0]?.url.includes('.bin:') ?? false)).toBe(true);
  });

  it('convert-mail-attachment-to-pdf returns an empty raw-bytes envelope for a plain-text attachment with no contentBytes', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'note.txt' }); // no contentBytes field
    const result = await pdfCmd().execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'a1' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { contentType: string; size: number; base64: string; note: string };
    expect(env.contentType).toBe('text/plain');
    expect(env.size).toBe(0); // empty contentBytes default '' → 0 bytes (mutant default is non-empty)
    expect(env.base64).toBe('');
    expect(env.note).toContain('note.txt');
  });

  it('convert-mail-attachment-to-pdf deletes the temp file AND the temp folder when the cleanup probe shows it empty', async () => {
    const deletes: string[] = [];
    await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('report.docx', { deletes, children: {} })), { messageId: 'm1', attachmentId: 'a1' });
    expect(deletes.some((u) => u.includes('/me/drive/items/temp-h1'))).toBe(true); // temp file always deleted
    // children body has no `value` → the `?? []` default makes it empty → folder
    // deleted. A non-empty default mutant would skip this delete.
    expect(deletes.some((u) => u.endsWith('/.ask-marcel-temp'))).toBe(true);
  });

  it('convert-mail-attachment-to-pdf propagates an upload (PUT) failure rather than mislabelling it as a missing id', async () => {
    const puts: PutRecord[] = [];
    const result = await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('report.docx', { puts, putStatus: 403 })), { messageId: 'm1', attachmentId: 'a1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.status).toBe(403); // mutant skips the guard → 500 'no driveItem id'
    expect(result.error.message).not.toContain('upload returned no driveItem id');
  });

  it('convert-mail-attachment-to-pdf propagates the initial attachment-fetch failure', async () => {
    const result = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => new Response('nope', { status: 404 })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.status).toBe(404); // mutant skips the guard → reads undefined.value → throws
  });

  it('convert-mail-attachment-to-pdf errs unconditionally for itemAttachment, missing @odata.type, and unknown type', async () => {
    const item = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ '@odata.type': '#microsoft.graph.itemAttachment', item: {} })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(item.ok).toBe(false);
    if (!item.ok) {
      expect(item.error.type).toBe('api_error');
      expect(item.error.message).toContain('convert-mail-attachment-to-markdown');
      if (item.error.type === 'api_error') expect(item.error.status).toBe(400);
    }
    const noType = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ name: 'x.docx' })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(noType.ok).toBe(false);
    if (!noType.ok) {
      expect(noType.error.type).toBe('api_error');
      expect(noType.error.message).toContain('missing @odata.type discriminator');
      if (noType.error.type === 'api_error') expect(noType.error.status).toBe(400);
    }
    const unknown = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ '@odata.type': '#microsoft.graph.weird' })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) {
      expect(unknown.error.type).toBe('api_error');
      expect(unknown.error.message).toContain('unsupported attachment type: #microsoft.graph.weird');
      if (unknown.error.type === 'api_error') expect(unknown.error.status).toBe(400);
    }
  });

  it('convert-mail-attachment-to-pdf errs unconditionally on referenceAttachment edge cases (missing/empty sourceUrl, resolved missing id, resolved missing driveId, image source) and on resolve / upload failures', async () => {
    // missing sourceUrl (undefined)
    const noUrl = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment' })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(noUrl.ok).toBe(false);
    if (!noUrl.ok) {
      expect(noUrl.error.type).toBe('api_error');
      expect(noUrl.error.message).toContain('referenceAttachment missing sourceUrl');
      if (noUrl.error.type === 'api_error') expect(noUrl.error.status).toBe(400);
    }

    // empty-string sourceUrl exercises the `=== ''` operand specifically
    const emptyUrl = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: '' })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(emptyUrl.ok).toBe(false);
    if (!emptyUrl.ok) {
      expect(emptyUrl.error.type).toBe('api_error');
      expect(emptyUrl.error.message).toContain('referenceAttachment missing sourceUrl');
    }

    // resolved driveItem with no id (driveId present → the itemId operand is the deciding one)
    const itemIdMissing: FetchFn = async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/q.docx' });
      return Response.json({ name: 'q.docx', parentReference: { driveId: 'd1' } });
    };
    const iim = await pdfCmd().execute(createGraphClient(fakeAuth(), itemIdMissing), { messageId: 'm1', attachmentId: 'a1' });
    expect(iim.ok).toBe(false);
    if (!iim.ok) {
      expect(iim.error.type).toBe('api_error');
      expect(iim.error.message).toContain('resolved driveItem missing id or driveId');
    }

    // resolved driveItem with no driveId (the first operand)
    const driveIdMissing: FetchFn = async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/q.docx' });
      return Response.json({ id: 'i1', name: 'q.docx' });
    };
    const dim = await pdfCmd().execute(createGraphClient(fakeAuth(), driveIdMissing), { messageId: 'm1', attachmentId: 'a1' });
    expect(dim.ok).toBe(false);
    if (!dim.ok) {
      expect(dim.error.type).toBe('api_error');
      expect(dim.error.message).toContain('resolved driveItem missing id or driveId');
      if (dim.error.type === 'api_error') expect(dim.error.status).toBe(500);
    }

    // resolved driveItem that is an image → 415 (reference-side image guard)
    const refImage: FetchFn = async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/photo.png' });
      return Response.json({ id: 'i1', name: 'photo.png', parentReference: { driveId: 'd1' } });
    };
    const ri = await pdfCmd().execute(createGraphClient(fakeAuth(), refImage), { messageId: 'm1', attachmentId: 'a1' });
    expect(ri.ok).toBe(false);
    if (!ri.ok) {
      expect(ri.error.type).toBe('api_error');
      expect(ri.error.message).toContain('png attachment is an image');
      if (ri.error.type === 'api_error') expect(ri.error.status).toBe(415);
    }

    // /shares resolve failure propagates
    const resolveFail: FetchFn = async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/q.docx' });
      return new Response('nope', { status: 404 });
    };
    const rf = await pdfCmd().execute(createGraphClient(fakeAuth(), resolveFail), { messageId: 'm1', attachmentId: 'a1' });
    expect(rf.ok).toBe(false);
    if (rf.ok) return;
    expect(rf.error.type).toBe('api_error');
    if (rf.error.type === 'api_error') expect(rf.error.status).toBe(404);
  });

  it('convert-mail-attachment-to-pdf errs unconditionally when the upload returns no driveItem id', async () => {
    const noUploadId: FetchFn = async (url, init) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'r.docx', contentBytes: btoa('hi') });
      if (init?.method === 'PUT') return Response.json({ name: 'no-id-returned' }); // upload response without `id`
      throw new Error(`unexpected ${url}`);
    };
    const uploadNoId = await pdfCmd().execute(createGraphClient(fakeAuth(), noUploadId), { messageId: 'm1', attachmentId: 'a1' });
    expect(uploadNoId.ok).toBe(false);
    if (!uploadNoId.ok) {
      expect(uploadNoId.error.type).toBe('api_error');
      expect(uploadNoId.error.message).toContain('upload returned no driveItem id');
      if (uploadNoId.error.type === 'api_error') expect(uploadNoId.error.status).toBe(500);
    }
  });

  it('convert-mail-attachment-to-pdf rejects a referenceAttachment with no sourceUrl', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment' });
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefBad' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('missing sourceUrl');
  });

  it('convert-mail-attachment-to-pdf rejects a referenceAttachment whose resolved driveItem lacks ids', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefBad2')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/x.docx' });
      }
      if (url.includes('/shares/u!')) return Response.json({ id: 'no-drive' });
      throw new Error('unexpected');
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefBad2' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('missing id or driveId');
  });

  it('convert-mail-attachment-to-pdf short-circuits a referenceAttachment to raw bytes for plain-text source', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefText')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/notes.txt' });
      }
      if (url.includes('/shares/u!')) return Response.json({ id: 'i-text', name: 'notes.txt', parentReference: { driveId: 'd1' } });
      if (url.endsWith('/drives/d1/items/i-text/content')) return new Response('plain', { status: 200, headers: { 'content-type': 'text/plain' } });
      throw new Error('unexpected');
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefText' });
    expect(result.ok).toBe(true);
  });

  it('convert-mail-attachment-to-pdf returns api_error when upload returns no driveItem id', async () => {
    // Distinct attachment id so the fake's URL-suffix check differs from the markdown variant's test above.
    const fetchFn: FetchFn = async (url, init) => {
      if (url.endsWith('/attachments/aFilePdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: btoa('docx-bytes') });
      }
      if (init?.method === 'PUT') return Response.json({ name: 'no-id' });
      throw new Error('unexpected');
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aFilePdf' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('upload returned no driveItem id');
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
