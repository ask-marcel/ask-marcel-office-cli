import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';

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
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined }),
  ...overrides,
});

describe('convertToMarkdown — orchestrate getBinary + optional 302 follow + image embedding + turndown', () => {
  it('converts inline HTML returned by getBinary directly', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'text/html', size: 24, text: '<h1>Q3 Budget</h1>' }),
    });
    const result = await convertToMarkdown(graph, '/drives/d1/items/i1/content?format=html');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('# Q3 Budget');
    }
  });

  it('follows a 302 downloadUrl via fetchUrl when getBinary returns it instead of inline bytes', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/_layouts/converted.html' }),
      fetchUrl: async (url) => {
        expect(url).toBe('https://contoso.sharepoint.com/_layouts/converted.html');
        return ok({ contentType: 'text/html', size: 13, text: '<p>follow</p>' });
      },
    });
    const result = await convertToMarkdown(graph, '/drives/d1/items/i1/content?format=html');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('follow');
    }
  });

  it('embeds inline images from the supplied attachments before running turndown', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'text/html', size: 60, text: '<p>Logo: <img src="cid:l1" alt="logo"></p>' }),
    });
    const result = await convertToMarkdown(graph, '/me/messages/m1?$select=body', [{ contentId: 'l1', contentType: 'image/png', contentBytes: 'iVBORw0=' }]);
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('data:image/png;base64,iVBORw0=');
      expect(v.text).not.toContain('cid:l1');
    }
  });

  it('propagates an err from getBinary unchanged (auth / api / network)', async () => {
    const graph = noopGraph({
      getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 404, message: 'not found' } }),
    });
    const result = await convertToMarkdown(graph, '/missing');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('api_error');
  });

  it('propagates an err from fetchUrl when following a 302 redirect fails', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/x.html' }),
      fetchUrl: async () => ({ ok: false, error: { type: 'network_error' as const, message: 'socket reset' } }),
    });
    const result = await convertToMarkdown(graph, '/drives/d1/items/i1/content?format=html');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toBe('socket reset');
    }
  });

  it('errors with a clear message when getBinary returns an unexpected envelope shape', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ surprise: true }),
    });
    const result = await convertToMarkdown(graph, '/some/path');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('unexpected response shape');
    }
  });

  it('also accepts the "redirect-followed-by-fetchUrl" path when fetchUrl returns the same envelope shape', async () => {
    // covers the case where fetchUrl returns text envelope (not direct text)
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/x.html' }),
      fetchUrl: async () => ok({ contentType: 'text/html', size: 5, text: '<p>x</p>' }),
    });
    const result = await convertToMarkdown(graph, '/whatever');
    expect(result.ok).toBe(true);
  });

  it('errors when fetchUrl returns an envelope without a text field (e.g. unexpected binary)', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/x.bin' }),
      fetchUrl: async () => ok({ contentType: 'application/octet-stream', size: 4, base64: 'AAAA' }),
    });
    const result = await convertToMarkdown(graph, '/whatever');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('missing text field');
    }
  });

  it('rewrites the Sandbox_InputFormatNotSupported error into a clear hint citing the documented HTML input set + *-as-pdf workaround', async () => {
    const graph = noopGraph({
      getBinary: async () =>
        err({
          type: 'api_error' as const,
          status: 406,
          message: 'Sandbox_InputFormatNotSupported: An exception occurred while executing within the Sandbox',
        }),
    });
    const result = await convertToMarkdown(graph, '/drives/d1/items/i1/content?format=html');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('loop, fluid, wbtx, whiteboard');
      expect(result.error.message).toContain('*-as-pdf');
    }
  });

  it('rewrites the Sandbox error when it surfaces from the fetchUrl 302-follow path (CDN refusal, not the initial Graph call)', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://francecentral1-mediap.svc.ms/transform/html?x=1' }),
      fetchUrl: async () =>
        err({
          type: 'api_error' as const,
          status: 406,
          message: 'Sandbox_InputFormatNotSupported: An exception occurred while executing within the Sandbox',
        }),
    });
    const result = await convertToMarkdown(graph, '/drives/d1/items/i1/content?format=html');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('loop, fluid, wbtx, whiteboard');
    }
  });
});
