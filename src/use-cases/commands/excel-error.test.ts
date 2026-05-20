import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { mapWacError, wrapExcelExecute } from './excel-error.ts';

describe('mapWacError', () => {
  it('rewrites the "Could not obtain a WAC access token" Graph error into a clear "not a workbook" hint', () => {
    const original: GraphError = { type: 'api_error', status: 403, message: 'AccessDenied: Could not obtain a WAC access token.' };
    const mapped = mapWacError(original);
    expect(mapped.type).toBe('api_error');
    if (mapped.type === 'api_error') {
      expect(mapped.status).toBe(403);
      expect(mapped.message).toContain('not an accessible Excel workbook');
      expect(mapped.message).toContain('get-drive-item');
    }
  });

  it('passes through non-api_error errors unchanged', () => {
    const original: GraphError = { type: 'auth_failed', message: 'Auth cancelled' };
    expect(mapWacError(original)).toEqual(original);
  });

  it('passes through api_errors that are not about WAC access tokens', () => {
    const original: GraphError = { type: 'api_error', status: 404, message: 'itemNotFound' };
    expect(mapWacError(original)).toEqual(original);
  });
});

describe('wrapExcelExecute', () => {
  const stubGraph: GraphClient = {
    get: async () => ok({}),
    post: async () => ok({}),
    getBinary: async () => ok({}),
    getElevated: async () => ok({}),
    teamsChat: async () => ok({}),
    getBinaryElevated: async () => ok({}),
    fetchUrl: async () => ok({}),
    put: async () => ok({}),
    delete: async () => ok({}),
    getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined }),
  };

  it('forwards the inner ok value unchanged', async () => {
    const wrapped = wrapExcelExecute(async () => ok({ value: [{ name: 'Sheet1' }] }));
    const result = await wrapped(stubGraph, {});
    expect(result).toEqual(ok({ value: [{ name: 'Sheet1' }] }));
  });

  it('rewrites the WAC error when the inner result is err', async () => {
    const wrapped = wrapExcelExecute(async () => err({ type: 'api_error', status: 403, message: 'AccessDenied: Could not obtain a WAC access token.' }));
    const result = await wrapped(stubGraph, {});
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('not an accessible Excel workbook');
  });

  it('passes through non-WAC errors unchanged', async () => {
    const original: GraphError = { type: 'api_error', status: 404, message: 'itemNotFound' };
    const wrapped = wrapExcelExecute(async () => err(original));
    const result = await wrapped(stubGraph, {});
    expect(result).toEqual(err(original));
  });
});
