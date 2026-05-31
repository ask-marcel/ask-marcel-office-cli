import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { buildSampleDocx } from '../../test-helpers/office-fixtures.ts';
import { execute } from './download-drive-item-version.ts';

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

// The historical-version endpoint MUST use the elevated token: getBinary (the
// non-elevated path) errors here, so any mutant that drops `elevated: true`
// flips to getBinary and fails — killing the elevation mutants.
const versionGraph = (handlers: { get?: (url: string) => Result<unknown, GraphError>; elevated?: (url: string) => Result<unknown, GraphError> }): GraphClient =>
  ({
    get: async (url: string) => handlers.get?.(url) ?? ok({}),
    post: async () => ok({}),
    getBinary: async () => ({ ok: false, error: { type: 'api_error', status: 403, message: 'non-elevated token rejected on historical version' } }),
    getElevated: async () => ok({}),
    teamsChat: async () => ok({}),
    teamsChatIc3: async () => ok({}),
    getBinaryElevated: async (url: string) => handlers.elevated?.(url) ?? ok({}),
    fetchUrl: async () => ok({}),
    put: async () => ok({}),
    delete: async () => ok({}),
    getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
  }) as GraphClient;

const params = { driveId: 'd1', itemId: 'i1', versionId: '2.0' };
const val = (r: Result<unknown, GraphError>): Record<string, unknown> => (r.ok ? (r.value as Record<string, unknown>) : {});

describe('download-drive-item-version', () => {
  it('returns a validation_error when a required field is missing', async () => {
    const result = await execute(versionGraph({}), { driveId: 'd1', itemId: 'i1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  it('format=original (default) returns the raw bytes via the elevated token', async () => {
    const graph = versionGraph({
      elevated: () => ok({ contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', size: 3, base64: toBase64(new Uint8Array([1, 2, 3])) }),
    });
    const result = await execute(graph, params);
    expect(result.ok).toBe(true);
    expect(typeof val(result).base64).toBe('string');
    expect(val(result).passthrough).toBeUndefined();
  });

  it('format=pdf on an Office source runs the conversion and (when Graph really returns a PDF) does not tag passthrough', async () => {
    const graph = versionGraph({
      get: () => ok({ name: 'report.docx' }),
      elevated: (url) => (url.includes('format=pdf') ? ok({ contentType: 'application/pdf', size: 4, base64: toBase64(new Uint8Array([37, 80, 68, 70])) }) : ok({})),
    });
    const result = await execute(graph, { ...params, format: 'pdf' });
    expect(result.ok).toBe(true);
    expect(val(result).contentType).toBe('application/pdf');
    expect(val(result).passthrough).toBeUndefined();
  });

  it('format=pdf tags passthrough + a note when Graph silently falls back to raw source bytes', async () => {
    const graph = versionGraph({
      get: () => ok({ name: 'report.docx' }),
      elevated: () => ok({ contentType: 'application/octet-stream', size: 3, base64: toBase64(new Uint8Array([1, 2, 3])) }),
    });
    const result = await execute(graph, { ...params, format: 'pdf' });
    expect(result.ok).toBe(true);
    expect(val(result).passthrough).toBe(true);
    expect(String(val(result).note)).toContain('format=pdf conversion was NOT applied');
  });

  it('format=pdf on a pdf source short-circuits to raw bytes with an "already PDF" note (no format=pdf round-trip)', async () => {
    const graph = versionGraph({
      get: () => ok({ name: 'scan.pdf' }),
      elevated: (url) =>
        url.includes('format=pdf') ? ok({ contentType: 'application/pdf' }) : ok({ contentType: 'application/pdf', size: 1, base64: toBase64(new Uint8Array([1])) }),
    });
    const result = await execute(graph, { ...params, format: 'pdf' });
    expect(result.ok).toBe(true);
    expect(val(result).passthrough).toBe(true);
    expect(String(val(result).note)).toContain('already PDF');
  });

  it('format=pdf on a plain-text source short-circuits to raw bytes with a plain-text note', async () => {
    const graph = versionGraph({
      get: () => ok({ name: 'notes.txt' }),
      elevated: () => ok({ contentType: 'text/plain', size: 5, base64: toBase64(new Uint8Array([104, 101, 108, 108, 111])) }),
    });
    const result = await execute(graph, { ...params, format: 'pdf' });
    expect(result.ok).toBe(true);
    expect(val(result).passthrough).toBe(true);
    expect(String(val(result).note)).toContain('plain-text');
  });

  it('format=markdown converts the docx body (elevated) and threads includeMetadata', async () => {
    const docx = await buildSampleDocx();
    const graph = versionGraph({
      get: () => ok({ name: 'report.docx' }),
      elevated: () => ok({ contentType: 'application/octet-stream', size: docx.byteLength, base64: toBase64(docx) }),
    });
    const plain = await execute(graph, { ...params, format: 'markdown' });
    expect(plain.ok).toBe(true);
    expect(String(val(plain).text)).toContain('# Sample Heading');
    expect(String(val(plain).text)).not.toContain('## DOCX metadata');
    const withMeta = await execute(graph, { ...params, format: 'markdown', includeMetadata: 'true' });
    expect(String(val(withMeta).text)).toContain('## DOCX metadata');
  });

  it('propagates the filename-metadata fetch error on the pdf path', async () => {
    const graph = versionGraph({ get: () => ({ ok: false, error: { type: 'api_error', status: 404, message: 'version meta gone' } }) });
    const result = await execute(graph, { ...params, format: 'pdf' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.message).toBe('version meta gone');
  });
});
