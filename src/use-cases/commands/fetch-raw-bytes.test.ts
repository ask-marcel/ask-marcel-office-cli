import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fetchRawBytes, inlineBinary } from './fetch-raw-bytes.ts';

const noopGraph = (overrides: Partial<GraphClient>): GraphClient => ({
  get: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ok({}),
  teamsChat: async () => ok({}),
  teamsChatIc3: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
  ...overrides,
});

const PDF_HEADER = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]); // %PDF-
const PDF_HEADER_B64 = 'JVBERi0=';

describe('fetchRawBytes', () => {
  it('returns the inline base64 bytes when Graph hands back a direct binary response (no redirect)', async () => {
    const graph = noopGraph({
      getBinary: async (path) => {
        expect(path).toBe('/drives/d1/items/i1/content');
        return ok({ contentType: 'application/pdf', size: 5, base64: PDF_HEADER_B64 });
      },
    });
    const result = await fetchRawBytes(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(true);
    if (result.ok) expect(Array.from(result.value)).toEqual(Array.from(PDF_HEADER));
  });

  it('follows the @microsoft.graph.downloadUrl redirect via graph.fetchUrl when present', async () => {
    const fetched: { url?: string } = {};
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/file.pdf' }),
      fetchUrl: async (url) => {
        fetched.url = url;
        return ok({ contentType: 'application/pdf', size: 5, base64: PDF_HEADER_B64 });
      },
    });
    const result = await fetchRawBytes(graph, '/drives/d1/items/i1/content?format=pdf');
    expect(result.ok).toBe(true);
    expect(fetched.url).toBe('https://contoso.sharepoint.com/cdn/file.pdf');
  });

  it('uses the elevated token path when opts.elevated is true (M365ChatClient identity for ODSP-gated bytes)', async () => {
    let basicCalled = 0;
    let elevatedCalled = 0;
    const graph = noopGraph({
      getBinary: async () => {
        basicCalled += 1;
        return ok({});
      },
      getBinaryElevated: async () => {
        elevatedCalled += 1;
        return ok({ contentType: 'application/pdf', size: 5, base64: PDF_HEADER_B64 });
      },
    });
    await fetchRawBytes(graph, '/drives/d1/items/i1/versions/3.0/content', { elevated: true });
    expect(elevatedCalled).toBe(1);
    expect(basicCalled).toBe(0);
  });

  it('decodes a text-content-type response into UTF-8 bytes (plain-text passthrough path)', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'text/plain', size: 5, text: 'hello' }),
    });
    const result = await fetchRawBytes(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(true);
    if (result.ok) expect(new TextDecoder().decode(result.value)).toBe('hello');
  });

  it('returns api_error when the Graph response carries neither base64 nor text nor a downloadUrl', async () => {
    const graph = noopGraph({ getBinary: async () => ok({ unexpected: 'shape' }) });
    const result = await fetchRawBytes(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('unexpected envelope');
    }
  });

  it('passes through the underlying GraphError when getBinary fails', async () => {
    const graph = noopGraph({ getBinary: async () => err({ type: 'auth_failed', message: 'Unauthorized' }) });
    const result = await fetchRawBytes(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_failed');
  });

  it('passes through the underlying GraphError when fetchUrl fails after a redirect', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/file.pdf' }),
      fetchUrl: async () => err({ type: 'network_error', message: 'CDN timeout' }),
    });
    const result = await fetchRawBytes(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('network_error');
  });
});

describe('inlineBinary', () => {
  it('returns { contentType, size, base64 } unchanged when Graph hands back inline base64 directly', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/pdf', size: 5, base64: PDF_HEADER_B64 }),
    });
    const result = await inlineBinary(graph, '/drives/d1/items/i1/content?format=pdf');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toEqual({ contentType: 'application/pdf', size: 5, base64: PDF_HEADER_B64 });
    }
  });

  it('follows the @microsoft.graph.downloadUrl redirect and returns the inline binary envelope from the CDN', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/deck.pdf' }),
      fetchUrl: async () => ok({ contentType: 'application/pdf', size: 12345, base64: PDF_HEADER_B64 }),
    });
    const result = await inlineBinary(graph, '/drives/d1/items/i1/content?format=pdf');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.contentType).toBe('application/pdf');
      expect(result.value.size).toBe(12345);
      expect(result.value.base64).toBe(PDF_HEADER_B64);
    }
  });

  it('uses getBinaryElevated for opts.elevated (historical-version PDF path)', async () => {
    let elevatedCalls = 0;
    const graph = noopGraph({
      getBinaryElevated: async () => {
        elevatedCalls += 1;
        return ok({ contentType: 'application/pdf', size: 5, base64: PDF_HEADER_B64 });
      },
    });
    await inlineBinary(graph, '/drives/d1/items/i1/versions/3.0/content?format=pdf', { elevated: true });
    expect(elevatedCalls).toBe(1);
  });

  it('falls back to application/octet-stream when contentType is missing on the upstream envelope', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ size: 5, base64: PDF_HEADER_B64 }),
    });
    const result = await inlineBinary(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.contentType).toBe('application/octet-stream');
  });

  it('re-encodes a text-content-type response as base64 so downstream consumers always get a uniform binary envelope', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'text/plain', size: 5, text: 'hello' }),
    });
    const result = await inlineBinary(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.contentType).toBe('text/plain');
      expect(result.value.base64).toBe(btoa('hello'));
    }
  });

  it('returns api_error when the upstream envelope has neither base64 nor text nor a downloadUrl', async () => {
    const graph = noopGraph({ getBinary: async () => ok({ unexpected: 'shape' }) });
    const result = await inlineBinary(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('unexpected envelope');
    }
  });

  it('passes through the underlying GraphError when getBinary fails', async () => {
    const graph = noopGraph({ getBinary: async () => err({ type: 'api_error', status: 404, message: 'item not found' }) });
    const result = await inlineBinary(graph, '/drives/d1/items/missing/content');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.status).toBe(404);
  });

  it('infers size from base64 length when the upstream envelope omits an explicit size field', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/pdf', base64: PDF_HEADER_B64 }),
    });
    const result = await inlineBinary(graph, '/drives/d1/items/i1/content');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.size).toBe(6); // floor(8 * 3 / 4)
  });
});

describe('tagPdfPassthrough', () => {
  it('passes through unchanged when the response IS application/pdf (the conversion succeeded — no warning needed)', async () => {
    const { tagPdfPassthrough } = await import('./fetch-raw-bytes.ts');
    const inner = ok({ contentType: 'application/pdf', size: 100, base64: 'JVBERi0=' });
    const tagged = tagPdfPassthrough(inner, 'deck.pptx');
    expect(tagged).toEqual(inner);
  });

  it('passes through err results unchanged (no point tagging a failure)', async () => {
    const { tagPdfPassthrough } = await import('./fetch-raw-bytes.ts');
    const inner = err({ type: 'api_error' as const, status: 500, message: 'boom' });
    const tagged = tagPdfPassthrough(inner, 'deck.pptx');
    expect(tagged).toEqual(inner);
  });

  it('attaches passthrough:true and a "save with source extension" note when the response is NOT application/pdf despite a format=pdf request (audit round-5 #2 — the silent raw-bytes fallback)', async () => {
    const { tagPdfPassthrough } = await import('./fetch-raw-bytes.ts');
    const inner = ok({ contentType: 'application/octet-stream', size: 980167, base64: 'rawpptx' });
    const tagged = tagPdfPassthrough(inner, 'roadmap26.pptx');
    expect(tagged.ok).toBe(true);
    if (tagged.ok) {
      const v = tagged.value as { contentType: string; passthrough: true; note: string };
      expect(v.contentType).toBe('application/octet-stream');
      expect(v.passthrough).toBe(true);
      expect(v.note).toContain('roadmap26.pptx');
      expect(v.note).toContain('format=pdf conversion was NOT applied');
      expect(v.note).toContain('save with the source extension');
    }
  });

  it('also tags when the response carries a non-pdf MIME like vnd.openxmlformats-officedocument.presentationml.presentation', async () => {
    const { tagPdfPassthrough } = await import('./fetch-raw-bytes.ts');
    const inner = ok({
      contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      size: 980167,
      base64: 'rawpptx',
    });
    const tagged = tagPdfPassthrough(inner, 'deck.pptx');
    expect(tagged.ok).toBe(true);
    if (tagged.ok) {
      const v = tagged.value as { passthrough: boolean; note: string };
      expect(v.passthrough).toBe(true);
      expect(v.note).toContain('vnd.openxmlformats-officedocument.presentationml.presentation');
    }
  });
});
