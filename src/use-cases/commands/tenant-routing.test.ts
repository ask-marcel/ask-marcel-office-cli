import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute as convertZip } from './convert-drive-item-zip-to-markdown.ts';
import { execute as asMarkdown } from './download-drive-item-as-markdown.ts';
import { execute as asPdf } from './download-drive-item-as-pdf.ts';
import { execute as downloadContent } from './download-drive-item-content.ts';
import { execute as extractImages } from './extract-drive-item-images.ts';

const PARTNER = tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04');
const ARGS = { driveId: 'b!x', itemId: '01ABC', tenantId: PARTNER };

/**
 * A partner-tenant file is unreadable on the home token — EVERY call, not just
 * the download. Several of these commands read the item's metadata first (to
 * detect a folder, or to learn the filename), and that read is the FIRST call:
 * leaving it on the home token fails with `invalidAudienceUri` before the bytes
 * are ever requested.
 *
 * That is not hypothetical. It shipped: `download-drive-item-as-pdf` and
 * `extract-drive-item-images` had their byte fetch routed and their metadata
 * fetch left behind, and the whole unit suite was green because nothing asserted
 * the HOME client is never touched. It took a live run to surface. These tests
 * make the home client throw, so any unrouted call fails loudly here instead.
 */
const guestOnlyGraph = (calls: string[]): GraphClient =>
  fakeGraphClient({
    get: (path: string) => {
      throw new Error(`home token used for a partner-tenant file: GET ${path}`);
    },
    getBinary: (path: string) => {
      throw new Error(`home token used for a partner-tenant file: GET ${path}`);
    },
    getBinaryElevated: (path: string) => {
      throw new Error(`elevated (home) token used for a partner-tenant file: GET ${path}`);
    },
    getGuest: async (path: string) => {
      calls.push(`guest:${path}`);
      return ok({ id: '01ABC', name: 'deck.pptx' });
    },
    getBinaryGuest: async (path: string) => {
      calls.push(`guestBinary:${path}`);
      return ok({ contentType: 'text/plain', size: 5, base64: btoa('hello') });
    },
  });

describe('reading a partner-tenant file end to end', () => {
  it('never touches the home token when downloading content from a partner tenant', async () => {
    const calls: string[] = [];

    const result = await downloadContent(guestOnlyGraph(calls), { ...ARGS });

    expect(result.ok).toBe(true);
    expect(calls.some((c) => c.startsWith('guest:/drives/b!x/items/01ABC'))).toBe(true);
    expect(calls.some((c) => c.startsWith('guestBinary:'))).toBe(true);
  });

  it('never touches the home token when converting a partner-tenant file to markdown', async () => {
    const calls: string[] = [];

    await asMarkdown(guestOnlyGraph(calls), { ...ARGS });

    expect(calls[0]).toBe('guest:/drives/b!x/items/01ABC');
  });

  // The two that shipped broken.
  it('never touches the home token when rendering a partner-tenant file to pdf', async () => {
    const calls: string[] = [];

    await asPdf(guestOnlyGraph(calls), { ...ARGS });

    expect(calls[0]).toBe('guest:/drives/b!x/items/01ABC');
  });

  it('never touches the home token when extracting images from a partner-tenant file', async () => {
    const calls: string[] = [];

    await extractImages(guestOnlyGraph(calls), { ...ARGS });

    expect(calls[0]).toBe('guest:/drives/b!x/items/01ABC');
  });

  it('never touches the home token when unpacking a partner-tenant zip', async () => {
    const calls: string[] = [];

    await convertZip(guestOnlyGraph(calls), { ...ARGS });

    expect(calls.some((c) => c.startsWith('guestBinary:'))).toBe(true);
  });
});
