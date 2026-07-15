import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './get-user.ts';

describe('get-user', () => {
  it('fetches the full profile via the elevated /users/{id} path for a GUID', async () => {
    let elevatedPath = '';
    const graph = fakeGraphClient({
      getElevated: async (p) => {
        elevatedPath = p;
        return ok({ id: 'aaaaaaaa-1111-2222-3333-444444444444', displayName: 'Alice Kim', jobTitle: 'Engineer', department: 'IS&T' });
      },
    });
    const result = await execute(graph, { userId: 'aaaaaaaa-1111-2222-3333-444444444444' });
    expect(elevatedPath).toContain('/users/aaaaaaaa-1111-2222-3333-444444444444');
    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { jobTitle?: string }).jobTitle).toBe('Engineer'); // org fields need the elevated token
  });

  it('fetches the full profile via the elevated path for a UPN/email, honouring --select', async () => {
    let elevatedPath = '';
    const graph = fakeGraphClient({
      getElevated: async (p) => {
        elevatedPath = p;
        return ok({ id: 'x', mail: 'alice@contoso.com' });
      },
    });
    await execute(graph, { userId: 'alice@contoso.com', select: 'id,displayName,mail' });
    expect(elevatedPath).toContain('/users/alice%40contoso.com'); // the @ is percent-encoded into the path segment
    expect(elevatedPath).toContain('$select=id%2CdisplayName%2Cmail'); // appendOData URL-encodes the commas
  });

  it('percent-encodes a guest UPN so the #EXT# fragment marker is not dropped from the path', async () => {
    let elevatedPath = '';
    const graph = fakeGraphClient({
      getElevated: async (p) => {
        elevatedPath = p;
        return ok({ id: 'g1' });
      },
    });
    await execute(graph, { userId: 'alice_contoso.com#EXT#@fabrikam.onmicrosoft.com' });
    expect(elevatedPath).toContain('/users/alice_contoso.com%23EXT%23%40fabrikam.onmicrosoft.com');
    expect(elevatedPath).not.toContain('#EXT#'); // a raw # would truncate the path at the URL-fragment boundary
  });

  it('searches the relevant-people graph for a bare name and returns candidate matches', async () => {
    let getPath = '';
    const graph = fakeGraphClient({
      get: async (p) => {
        getPath = p;
        return ok({
          value: [{ id: 'u1', displayName: 'Weilai Wang', jobTitle: 'Senior IT Manager', department: 'IS&T', scoredEmailAddresses: [{ address: 'weilai.wang@x.com' }] }],
        });
      },
    });
    const result = await execute(graph, { userId: 'Weilai Wang' });
    expect(getPath).toContain('/me/people?$search=');
    expect(getPath).toContain('Weilai Wang');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { query: string; matches: ReadonlyArray<Record<string, string | undefined>> };
      expect(v.query).toBe('Weilai Wang');
      expect(v.matches).toEqual([{ id: 'u1', displayName: 'Weilai Wang', mail: 'weilai.wang@x.com', jobTitle: 'Senior IT Manager', department: 'IS&T' }]);
    }
  });

  it('returns empty matches when the name search finds nobody in the relevant-people graph', async () => {
    const graph = fakeGraphClient({ get: async () => ok({ value: [] }) });
    const result = await execute(graph, { userId: 'Nobody Here' });
    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { matches: unknown[] }).matches).toEqual([]);
  });

  it('passes through the elevated fail-fast for an id when the elevated token is cold', async () => {
    const graph = fakeGraphClient({
      getElevated: async () => ({
        ok: false as const,
        error: { type: 'api_error' as const, status: 401, code: 'secondary_token_unavailable', message: 'run `ask-marcel-office login`' },
      }),
    });
    const result = await execute(graph, { userId: 'alice@contoso.com' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.code).toBe('secondary_token_unavailable');
  });

  it('rejects a missing user-id via Zod', async () => {
    const result = await execute(fakeGraphClient({}), {});
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('propagates a People API error on the name path', async () => {
    const graph = fakeGraphClient({ get: async () => ({ ok: false as const, error: { type: 'api_error' as const, status: 500, message: 'boom' } }) });
    const result = await execute(graph, { userId: 'Someone Unknown' });
    expect(result.ok).toBe(false);
  });

  it('routes a string that only contains a GUID (with surrounding junk) to the People search, never /users', async () => {
    const calls: string[] = [];
    const graph = fakeGraphClient({
      get: async () => {
        calls.push('people');
        return ok({ value: [] });
      },
      getElevated: async () => {
        calls.push('elevated');
        return ok({});
      },
    });
    await execute(graph, { userId: 'aaaaaaaa-1111-2222-3333-444444444444-trailing' }); // GUID + suffix
    await execute(graph, { userId: 'leading-aaaaaaaa-1111-2222-3333-444444444444' }); // prefix + GUID
    expect(calls).toEqual(['people', 'people']); // the anchors reject a partial GUID match
  });

  it('strips embedded quotes from a name so they cannot break out of the $search literal', async () => {
    let getPath = '';
    const graph = fakeGraphClient({
      get: async (p) => {
        getPath = p;
        return ok({ value: [] });
      },
    });
    await execute(graph, { userId: 'Ali"ce' });
    expect(getPath).toContain('$search="Alice"'); // the inner quote is removed, not left in or replaced
    expect(getPath).not.toContain('Ali"ce');
  });

  it('drops non-string People fields (a numeric id) to undefined rather than leaking the raw value', async () => {
    const graph = fakeGraphClient({ get: async () => ok({ value: [{ id: 42, displayName: 'Weird Shape' }] }) });
    const result = await execute(graph, { userId: 'Weird Shape' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const [match] = (result.value as { matches: ReadonlyArray<Record<string, string | undefined>> }).matches;
      expect(match?.id).toBeUndefined(); // 42 is not a string, so it is dropped
      expect(match?.displayName).toBe('Weird Shape');
    }
  });

  it('yields an undefined mail when a candidate has an empty scoredEmailAddresses array', async () => {
    const graph = fakeGraphClient({ get: async () => ok({ value: [{ id: 'u9', displayName: 'No Email', scoredEmailAddresses: [] }] }) });
    const result = await execute(graph, { userId: 'No Email' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const [match] = (result.value as { matches: ReadonlyArray<Record<string, string | undefined>> }).matches;
      expect(match?.mail).toBeUndefined(); // no [0] element to read .address off
    }
  });

  it('returns empty matches when the People response carries no value array', async () => {
    const graph = fakeGraphClient({ get: async () => ok({}) }); // no `value` key at all
    const result = await execute(graph, { userId: 'Ghost' });
    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { matches: unknown[] }).matches).toEqual([]);
  });
});
