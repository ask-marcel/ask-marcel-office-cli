import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleDocx, buildSampleMsg, buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { execute } from './download-drive-item-as-markdown.ts';

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

const graphWith = (handlers: { get?: (url: string) => Result<unknown, GraphError>; getBinary?: () => Result<unknown, GraphError> }): GraphClient =>
  fakeGraphClient({
    get: async (url: string) => handlers.get?.(url) ?? ok({}),
    getBinary: async () => handlers.getBinary?.() ?? ok({}),
  });

const asText = (r: Result<unknown, GraphError>): string => (r.ok ? ((r.value as { text?: string }).text ?? '') : '');
const params = { driveId: 'd1', itemId: 'i1' };

describe('download-drive-item-as-markdown', () => {
  it('returns a validation_error when driveId/itemId are missing', async () => {
    const result = await execute(graphWith({}), {});
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  it('propagates the item-metadata fetch error before converting', async () => {
    const graph = graphWith({ get: () => ({ ok: false, error: { type: 'api_error', status: 404, message: 'item gone' } }) });
    const result = await execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.message).toBe('item gone');
  });

  it('converts the docx body and omits the metadata block by default (includeMetadata not threaded as true)', async () => {
    const docx = await buildSampleDocx();
    const graph = graphWith({ get: () => ok({ name: 'r.docx' }), getBinary: () => ok({ contentType: 'application/octet-stream', size: docx.byteLength, base64: toBase64(docx) }) });
    const text = asText(await execute(graph, params));
    expect(text).toContain('# Sample Heading');
    expect(text).not.toContain('## DOCX metadata');
  });

  it('appends the DOCX metadata block only when includeMetadata is "true" (threads the flag through officeToMarkdown)', async () => {
    const docx = await buildSampleDocx();
    const graph = graphWith({ get: () => ok({ name: 'r.docx' }), getBinary: () => ok({ contentType: 'application/octet-stream', size: docx.byteLength, base64: toBase64(docx) }) });
    expect(asText(await execute(graph, { ...params, includeMetadata: 'true' }))).toContain('## DOCX metadata');
    expect(asText(await execute(graph, { ...params, includeMetadata: 'false' }))).not.toContain('## DOCX metadata');
  });

  it('threads --max-cells through to the xlsx converter so an oversized sheet is truncated to a hint instead of the full table', async () => {
    const xlsx = buildSampleXlsx();
    const graph = graphWith({
      get: () => ok({ name: 'big.xlsx' }),
      getBinary: () => ok({ contentType: 'application/octet-stream', size: xlsx.byteLength, base64: toBase64(xlsx) }),
    });
    const text = asText(await execute(graph, { ...params, maxCells: '4' }));
    expect(text).toContain('## Sheet1');
    expect(text).not.toContain('Alice');
    expect(text).toContain('get-excel-range');
  });
  it('embeds the image bytes as a data: URI only when --inline-images is "true"', async () => {
    const docx = await buildSampleDocx();
    const graph = graphWith({ get: () => ok({ name: 'r.docx' }), getBinary: () => ok({ contentType: 'application/octet-stream', size: docx.byteLength, base64: toBase64(docx) }) });
    expect(asText(await execute(graph, { ...params, inlineImages: 'true' }))).toContain('data:image/png;base64,');
  });

  it('leaves the images out of the markdown when --inline-images is absent, the default since the 2.2.0 flip', async () => {
    const docx = await buildSampleDocx();
    const graph = graphWith({ get: () => ok({ name: 'r.docx' }), getBinary: () => ok({ contentType: 'application/octet-stream', size: docx.byteLength, base64: toBase64(docx) }) });
    expect(asText(await execute(graph, params))).not.toContain('data:image/png;base64,');
    expect(asText(await execute(graph, { ...params, inlineImages: 'false' }))).not.toContain('data:image/png;base64,');
  });

  it('renders the whole sheet when --max-cells is absent, so the cap is opt-in rather than a silent default', async () => {
    const xlsx = buildSampleXlsx();
    const graph = graphWith({
      get: () => ok({ name: 'big.xlsx' }),
      getBinary: () => ok({ contentType: 'application/octet-stream', size: xlsx.byteLength, base64: toBase64(xlsx) }),
    });
    const text = asText(await execute(graph, params));
    expect(text).toContain('Alice');
    expect(text).not.toContain('get-excel-range');
  });

  it('refuses --max-cells "0" as not a positive integer, naming the constraint', async () => {
    const result = await execute(graphWith({}), { ...params, maxCells: '0' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(result.error.message).toContain('must be a positive integer');
  });

  it('refuses a --max-cells that only starts with a digit, so "1a" cannot reach Number() as NaN', async () => {
    const result = await execute(graphWith({}), { ...params, maxCells: '1a' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(result.error.message).toContain('must be a positive integer');
  });

  it('refuses a malformed --tenant-id before any Graph call, so a guest read never leaves with a bad audience', async () => {
    let called = false;
    const graph = graphWith({
      get: () => {
        called = true;
        return ok({ name: 'r.docx' });
      },
    });
    const result = await execute(graph, { ...params, tenantId: 'not-a-guid' });
    expect(result.ok).toBe(false);
    expect(called).toBe(false);
  });

  it('accepts --keep-quoted on a drive .msg and threads it to the renderer (the vendored fixture has no quoted chain, so both modes render the same body)', async () => {
    const msg = await buildSampleMsg();
    const graph = graphWith({ get: () => ok({ name: 'mail.msg' }), getBinary: () => ok({ contentType: 'application/octet-stream', size: msg.byteLength, base64: toBase64(msg) }) });
    const stripped = asText(await execute(graph, params));
    const kept = asText(await execute(graph, { ...params, keepQuoted: 'true' }));
    expect(stripped).toContain('Please find the quarterly figures attached.');
    expect(kept).toBe(stripped);
  });
});
