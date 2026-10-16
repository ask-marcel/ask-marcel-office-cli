import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import type { GraphError } from '../../infra/graph-request.ts';
import { fakeReadGraph } from '../../test-helpers/read-graph-fake.ts';
import { fakeWriteGraph } from '../../test-helpers/write-graph-fake.ts';
import { base64ToBytes, fetchRawBytes, inlineBasicBinary, inlineBinary, tagPdfPassthrough } from './fetch-raw-bytes.ts';

/*
 * The byte helpers read and write both use (core, package split phase 3): each
 * envelope Graph answers a file read with, inline or behind a CDN redirect,
 * becomes the same bytes.
 */

const PDF_HEADER_B64 = 'JVBERi0=';
const EMPTY_ENVELOPE: Result<never, GraphError> = err({
  type: 'api_error',
  status: 500,
  message: 'unexpected envelope: response had no @microsoft.graph.downloadUrl, no base64 bytes, and no text body',
});
const REFUSED: Result<never, GraphError> = err({ type: 'api_error', status: 404, message: 'itemNotFound: Gone.', code: 'itemNotFound' });
const CDN = 'https://contoso.sharepoint.com/cdn/f';

describe('the bytes of a file', () => {
  it('decode an inline base64 or text answer, and refuse an answer with neither', async () => {
    expect(await fetchRawBytes(fakeReadGraph({ getBinary: async () => ok({ base64: PDF_HEADER_B64 }) }), '/f')).toEqual(ok(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d])));
    expect(await fetchRawBytes(fakeReadGraph({ getBinary: async () => ok({ text: 'hé' }) }), '/f')).toEqual(ok(new TextEncoder().encode('hé')));
    expect(await fetchRawBytes(fakeReadGraph({ getBinary: async () => ok({ contentType: 'application/pdf' }) }), '/f')).toEqual(EMPTY_ENVELOPE);
  });

  it('follow a CDN redirect, and report a refused redirect or a refused read as it is', async () => {
    const followed: string[] = [];
    const redirecting = fakeReadGraph({
      getBinaryElevated: async () => ok({ '@microsoft.graph.downloadUrl': CDN }),
      fetchUrl: async (url) => {
        followed.push(url);
        return ok({ base64: PDF_HEADER_B64 });
      },
    });
    expect(await fetchRawBytes(redirecting, '/f', { elevated: true })).toEqual(ok(base64ToBytes(PDF_HEADER_B64)));
    expect(followed).toEqual([CDN]);
    expect(await fetchRawBytes(fakeReadGraph({ getBinary: async () => ok({ '@microsoft.graph.downloadUrl': CDN }), fetchUrl: async () => REFUSED }), '/f')).toEqual(REFUSED);
    expect(await fetchRawBytes(fakeReadGraph({ getBinary: async () => REFUSED }), '/f')).toEqual(REFUSED);
  });

  it('are read on the guest token of a named partner tenant, even when the elevated token is asked for too', async () => {
    const tiers: string[] = [];
    const graph = fakeReadGraph({
      getBinary: async () => (tiers.push('basic'), ok({ text: 'x' })),
      getBinaryElevated: async () => (tiers.push('elevated'), ok({ text: 'x' })),
      getBinaryGuest: async (_path, tenant) => (tiers.push(`guest ${tenant}`), ok({ text: 'x' })),
    });
    await fetchRawBytes(graph, '/f', { elevated: true, tenantId: tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04') });
    await fetchRawBytes(graph, '/f', { elevated: false });
    expect(tiers).toEqual(['guest 6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04', 'basic']);
  });
});

describe('an inlined file', () => {
  it('keeps the size and content type Graph declared', async () => {
    const graph = fakeReadGraph({ getBinary: async () => ok({ contentType: 'application/pdf', size: 99, base64: PDF_HEADER_B64 }) });
    expect(await inlineBinary(graph, '/f')).toEqual(ok({ contentType: 'application/pdf', size: 99, base64: PDF_HEADER_B64 }));
  });

  it('works the size out from the base64 or the UTF-8 text when Graph declared none, and names plain bytes when it gave no type', async () => {
    expect(await inlineBinary(fakeReadGraph({ getBinary: async () => ok({ base64: PDF_HEADER_B64 }) }), '/f')).toEqual(
      ok({ contentType: 'application/octet-stream', size: 6, base64: PDF_HEADER_B64 })
    );
    expect(await inlineBinary(fakeReadGraph({ getBinary: async () => ok({ contentType: 'text/plain', text: 'hé' }) }), '/f')).toEqual(
      ok({ contentType: 'text/plain', size: 3, base64: 'aMOp' })
    );
    expect(await inlineBinary(fakeReadGraph({ getBinary: async () => ok({ contentType: 'text/plain', size: 7, text: 'hé' }) }), '/f')).toEqual(
      ok({ contentType: 'text/plain', size: 7, base64: 'aMOp' })
    );
  });

  it('follows a CDN redirect, and reports a refused redirect, a refused read and an empty answer', async () => {
    const graph = fakeReadGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': CDN }),
      fetchUrl: async () => ok({ contentType: 'application/pdf', base64: PDF_HEADER_B64 }),
    });
    expect(await inlineBinary(graph, '/f')).toEqual(ok({ contentType: 'application/pdf', size: 6, base64: PDF_HEADER_B64 }));
    expect(await inlineBinary(fakeReadGraph({ getBinary: async () => ok({ '@microsoft.graph.downloadUrl': CDN }), fetchUrl: async () => REFUSED }), '/f')).toEqual(REFUSED);
    expect(await inlineBinary(fakeReadGraph({ getBinary: async () => REFUSED }), '/f')).toEqual(REFUSED);
    expect(await inlineBinary(fakeReadGraph({ getBinary: async () => ok({ size: 3 }) }), '/f')).toEqual(EMPTY_ENVELOPE);
  });

  it('is read on the basic token of the write graph too', async () => {
    const graph = fakeWriteGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': CDN }),
      fetchUrl: async () => ok({ contentType: 'application/rtf', text: '{\\rtf}' }),
    });
    expect(await inlineBasicBinary(graph, '/f')).toEqual(ok({ contentType: 'application/rtf', size: 6, base64: 'e1xydGZ9' }));
  });
});

describe('a PDF conversion Graph did not apply', () => {
  it('is tagged as the raw source with a note, while a real PDF and a failure pass as they are', () => {
    const pdf = ok({ contentType: 'application/pdf; charset=binary', size: 6, base64: PDF_HEADER_B64 });
    expect(tagPdfPassthrough(pdf, 'version 3')).toEqual(pdf);
    expect(tagPdfPassthrough(REFUSED, 'version 3')).toEqual(REFUSED);
    expect(tagPdfPassthrough(ok({ contentType: 'application/octet-stream', size: 6, base64: PDF_HEADER_B64 }), 'version 3 of deck.pptx')).toEqual(
      ok({
        contentType: 'application/octet-stream',
        size: 6,
        base64: PDF_HEADER_B64,
        passthrough: true,
        note: 'Graph returned `application/octet-stream` for version 3 of deck.pptx — format=pdf conversion was NOT applied (likely no Office Online runtime for this version/format on this tenant). The bytes are the raw source; save with the source extension, not .pdf.',
      })
    );
  });
});
