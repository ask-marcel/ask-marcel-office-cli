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

  it('injects a default $select with department on the direct id path when --select is absent (the help text promises the full org card)', async () => {
    let elevatedPath = '';
    const graph = fakeGraphClient({
      getElevated: async (p) => {
        elevatedPath = p;
        return ok({ id: 'x' });
      },
    });
    await execute(graph, { userId: 'aaaaaaaa-1111-2222-3333-444444444444' });
    expect(elevatedPath).toBe(
      '/users/aaaaaaaa-1111-2222-3333-444444444444?$select=id%2CdisplayName%2CuserPrincipalName%2Cmail%2CjobTitle%2Cdepartment%2CofficeLocation%2CbusinessPhones%2CmobilePhone'
    );
  });

  it('injects the same default $select on the mail-eq fallback path', async () => {
    let filterPath = '';
    const graph = fakeGraphClient({
      getElevated: async (p) => {
        if (p.includes('$filter=mail eq')) {
          filterPath = p;
          return ok({ value: [{ id: 'g1' }] });
        }
        return { ok: false as const, error: { type: 'api_error' as const, status: 404, code: 'Request_ResourceNotFound', message: 'missing' } };
      },
    });
    await execute(graph, { userId: 'guest@home.com' });
    expect(filterPath).toContain('$select=id%2CdisplayName%2CuserPrincipalName%2Cmail%2CjobTitle%2Cdepartment%2CofficeLocation%2CbusinessPhones%2CmobilePhone');
  });

  it('rejects a People-API contact id (base64-ish non-GUID candidate id) with a mail re-query remedy instead of silently name-searching it into empty matches', async () => {
    const calls: string[] = [];
    const graph = fakeGraphClient({
      get: async (p) => {
        calls.push(p);
        return ok({ value: [] });
      },
      getElevated: async (p) => {
        calls.push(p);
        return ok({});
      },
    });
    const result = await execute(graph, { userId: 'uBT42cGHMke_08PinHqwRg==' });
    expect(calls).toEqual([]); // no Graph round-trip: the id is unresolvable by design
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      expect(result.error.message).toContain('People-API contact id');
      expect(result.error.message).toContain('mail');
      expect(result.error.message).toContain('get-user --user-id');
    }
  });

  it('still routes ordinary display names to the People search (short, hyphenated, CJK, spaced, alphanumeric team names)', async () => {
    const searched: string[] = [];
    const graph = fakeGraphClient({
      get: async (p) => {
        searched.push(p);
        return ok({ value: [] });
      },
    });
    for (const name of ['Bob', 'Jean-Pierre', '王伟', 'Margaret Ma']) {
      const result = await execute(graph, { userId: name });
      expect(result.ok).toBe(true);
    }
    expect(searched.length).toBe(4);
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
    // Exact wire pin: quoted + percent-encoded (a raw space or & would
    // corrupt the query string; graph-client concatenates paths verbatim).
    expect(getPath).toBe('/me/people?$search=%22Weilai%20Wang%22');
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
    expect(getPath).toBe('/me/people?$search=%22Alice%22'); // the inner quote is removed, not left in or replaced
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

  const notFound = { ok: false as const, error: { type: 'api_error' as const, status: 404, code: 'Request_ResourceNotFound', message: 'Resource does not exist' } };

  it('falls back to a mail-eq filter when an email 404s on the direct path (a guest whose mail differs from their UPN)', async () => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      getElevated: async (p) => {
        paths.push(p);
        if (p.includes('$filter=mail eq')) return ok({ value: [{ id: 'g1', displayName: 'Robin Chen', mail: 'robin.chen@fabrikam.com' }] });
        return notFound; // the direct /users/{email} lookup misses (email is the mail, not the UPN)
      },
    });
    const result = await execute(graph, { userId: 'robin.chen@fabrikam.com' });
    expect(paths[0]).toContain('/users/robin.chen%40fabrikam.com'); // direct path, @ percent-encoded
    expect(paths[1]).toContain("/users?$filter=mail eq 'robin.chen@fabrikam.com'"); // fallback filter
    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { id: string }).id).toBe('g1'); // the matched user's full profile
  });

  it('carries --select through to the mail-eq fallback query', async () => {
    let filterPath = '';
    const graph = fakeGraphClient({
      getElevated: async (p) => {
        if (p.includes('$filter=mail eq')) {
          filterPath = p;
          return ok({ value: [{ id: 'g1' }] });
        }
        return notFound;
      },
    });
    await execute(graph, { userId: 'guest@home.com', select: 'id,displayName,mail' });
    expect(filterPath).toContain('$select=id%2CdisplayName%2Cmail'); // appended alongside the filter
  });

  it('surfaces the original 404 when the mail-eq fallback finds nobody', async () => {
    const graph = fakeGraphClient({
      getElevated: async (p) => (p.includes('$filter=mail eq') ? ok({ value: [] }) : notFound),
    });
    const result = await execute(graph, { userId: 'ghost@nowhere.com' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.status).toBe(404);
  });

  it('does not fall back when the direct path already resolves the email (mail == UPN)', async () => {
    let calls = 0;
    const graph = fakeGraphClient({
      getElevated: async () => {
        calls += 1;
        return ok({ id: 'u1', userPrincipalName: 'alice@contoso.com' });
      },
    });
    const result = await execute(graph, { userId: 'alice@contoso.com' });
    expect(calls).toBe(1); // no mail-filter round-trip when the direct GET succeeds
    expect(result.ok).toBe(true);
  });

  it('does not fall back for a GUID that 404s (a GUID is never an email)', async () => {
    let calls = 0;
    const graph = fakeGraphClient({
      getElevated: async () => {
        calls += 1;
        return notFound;
      },
    });
    const result = await execute(graph, { userId: 'aaaaaaaa-1111-2222-3333-444444444444' });
    expect(calls).toBe(1); // no `@`, so no mail fallback
    expect(result.ok).toBe(false);
  });

  it('does not fall back when an email fails with a non-404 (e.g. the cold-elevated fail-fast)', async () => {
    let calls = 0;
    const graph = fakeGraphClient({
      getElevated: async () => {
        calls += 1;
        return { ok: false as const, error: { type: 'api_error' as const, status: 401, code: 'secondary_token_unavailable', message: 'run `ask-marcel-office login --force`' } };
      },
    });
    const result = await execute(graph, { userId: 'alice@contoso.com' });
    expect(calls).toBe(1); // fallback only on a genuine 404, not on auth failures
    if (!result.ok && result.error.type === 'api_error') expect(result.error.code).toBe('secondary_token_unavailable');
  });

  it('surfaces the original 404 when the mail-eq fallback query itself errors', async () => {
    const graph = fakeGraphClient({
      getElevated: async (p) => (p.includes('$filter=mail eq') ? { ok: false as const, error: { type: 'api_error' as const, status: 500, message: 'filter boom' } } : notFound),
    });
    const result = await execute(graph, { userId: 'guest@home.com' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.status).toBe(404); // the direct 404, not the fallback's 500
  });

  it('doubles a single quote in the mail-eq filter (OData string-literal escaping)', async () => {
    let filterPath = '';
    const graph = fakeGraphClient({
      getElevated: async (p) => {
        if (p.includes('$filter=mail eq')) {
          filterPath = p;
          return ok({ value: [{ id: 'q1' }] });
        }
        return notFound;
      },
    });
    await execute(graph, { userId: "o'brien@x.com" });
    expect(filterPath).toContain("mail eq 'o''brien@x.com'"); // the ' is doubled, not left raw (prevents breaking the literal)
  });
});
