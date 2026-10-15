import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import { fetchRawBytes, inlineBinary } from './fetch-raw-bytes.ts';

// The byte fetch helpers are core: they take only the binary members of a
// graph, never a whole ReadGraph, so a package with another graph can call them.
const binaryMembersOnly = {
  getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/basic.txt' }),
  getBinaryElevated: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/elevated.txt' }),
  getBinaryGuest: async () => ok({ '@microsoft.graph.downloadUrl': 'https://partner.sharepoint.com/cdn/guest.txt' }),
  fetchUrl: async (url: string) => ok({ contentType: 'text/plain', text: url.split('/').at(-1) ?? '' }),
};

describe('the byte fetch helpers on a graph that holds only the binary members', () => {
  it('downloads a file on the basic token through its CDN redirect', async () => {
    const result = await fetchRawBytes(binaryMembersOnly, '/drives/d1/items/i1/content');
    expect(result).toEqual(ok(new TextEncoder().encode('basic.txt')));
  });

  it('inlines a file of a partner tenant on its guest token', async () => {
    const result = await inlineBinary(binaryMembersOnly, '/drives/d1/items/i1/content', { tenantId: tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04') });
    expect(result).toEqual(ok({ contentType: 'text/plain', size: 9, base64: 'Z3Vlc3QudHh0' }));
  });
});
