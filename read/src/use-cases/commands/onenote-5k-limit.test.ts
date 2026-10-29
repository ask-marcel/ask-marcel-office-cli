import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { createOnenote5kExecute } from './onenote-5k-limit.ts';

const LIMIT_SUMMARY =
  'OneNote read blocked: tenant has hit the 5000-item limit per document library (Graph error 10008). This is a SharePoint-side configuration issue, not a CLI bug — raise with the M365 admin to split the library or archive content. The CLI cannot work around it.';

describe('a OneNote read over the 5000-item limit', () => {
  const graph: GraphClient = fakeGraphClient();
  const failingWith = (error: GraphError): ReturnType<typeof createOnenote5kExecute> => createOnenote5kExecute(async () => err(error));

  it('forwards the inner ok value unchanged', async () => {
    const wrapped = createOnenote5kExecute(async () => ok({ value: [{ id: 'nb1' }] }));
    expect(await wrapped(graph, {})).toEqual(ok({ value: [{ id: 'nb1' }] }));
  });

  it('passes the parameters through to the inner command', async () => {
    const seen: Array<Record<string, string>> = [];
    const wrapped = createOnenote5kExecute(async (_graph, params) => {
      seen.push(params);
      return ok({});
    });
    await wrapped(graph, { siteId: 's1' });
    expect(seen).toEqual([{ siteId: 's1' }]);
  });

  it('rewrites the Graph error code 10008 into a one-line summary, keeping the status', async () => {
    const result = await failingWith({ type: 'api_error', status: 507, message: 'A long blog post.', code: '10008' })(graph, {});
    expect(result).toEqual(err({ type: 'api_error', status: 507, message: LIMIT_SUMMARY, code: 'cli_rewrite_onenote_5k_limit' }));
  });

  it('rewrites an error whose message names 10008 when the code does not', async () => {
    const result = await failingWith({ type: 'api_error', status: 400, message: 'Error 10008: too many items.', code: 'other' })(graph, {});
    expect(result).toEqual(err({ type: 'api_error', status: 400, message: LIMIT_SUMMARY, code: 'cli_rewrite_onenote_5k_limit' }));
  });

  it('passes any other API error through unchanged', async () => {
    const other: GraphError = { type: 'api_error', status: 404, message: 'Not found.', code: 'itemNotFound' };
    expect(await failingWith(other)(graph, {})).toEqual(err(other));
  });

  it('passes an error of another type through unchanged, even when it names 10008', async () => {
    const other: GraphError = { type: 'auth_failed', message: 'Code 10008 seen while signing in.' };
    expect(await failingWith(other)(graph, {})).toEqual(err(other));
  });
});
