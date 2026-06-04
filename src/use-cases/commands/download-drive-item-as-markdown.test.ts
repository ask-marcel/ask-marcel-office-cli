import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { buildSampleDocx, buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { execute } from './download-drive-item-as-markdown.ts';

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

const graphWith = (handlers: { get?: (url: string) => Result<unknown, GraphError>; getBinary?: () => Result<unknown, GraphError> }): GraphClient =>
  ({
    get: async (url: string) => handlers.get?.(url) ?? ok({}),
    post: async () => ok({}),
    getBinary: async () => handlers.getBinary?.() ?? ok({}),
    getElevated: async () => ok({}),
    teamsChat: async () => ok({}),
    teamsChatIc3: async () => ok({}),
    getBinaryElevated: async () => ok({}),
    fetchUrl: async () => ok({}),
    put: async () => ok({}),
    delete: async () => ok({}),
    getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
  }) as GraphClient;

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
});
