import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './resolve-drive-share-link.ts';

const PARTNER = tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04');

const driveItem = {
  id: '01ABCDEF',
  name: 'Q3 plan.pptx',
  size: 5347414,
  webUrl: 'https://fabrikam.sharepoint.com/:p:/s/Strategy/EaB1cD',
  lastModifiedDateTime: '2026-07-01T09:12:00Z',
  parentReference: { driveId: 'b!driveId' },
};

// The exact shape Graph answers with when a home-tenant token is pointed at a
// tenant the user is only a guest in: `00000003-0000-0ff1-ce00-000000000000` is
// SharePoint Online's app id, and home Graph cannot mint a token for a foreign
// tenant's SharePoint. Verified live 2026-07-16.
const foreignTenantAudienceError = err({
  type: 'api_error' as const,
  status: 401,
  message: "Invalid audience Uri '00000003-0000-0ff1-ce00-000000000000'.",
  code: 'invalidAudienceUri',
});

describe('resolving a sharing link to the file it points at', () => {
  it("resolves a link in the signed-in user's own tenant on the home token, with no tenant to pass on", async () => {
    const graph = fakeGraphClient({ get: async () => ok(driveItem) });

    const result = await execute(graph, { url: 'https://fabrikam.sharepoint.com/:p:/s/Strategy/EaB1cD?e=xyz' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as Record<string, unknown>;
    expect(value['driveId']).toBe('b!driveId');
    expect(value['itemId']).toBe('01ABCDEF');
    // Absent, deliberately: the home token reached it, so the caller needs no
    // `--tenant-id` on the follow-up. Presence of the field IS the signal.
    expect(value['tenantId']).toBeUndefined();
  });

  // The command holds no home-tenant id to compare the URL against, so it cannot
  // know a link is foreign up front. The audience error IS the signal: try home,
  // and only then discover + retry. Free on the home path, self-healing on the
  // foreign one.
  it('resolves a partner-tenant link by identifying the tenant and retrying as a guest', async () => {
    const calls: string[] = [];
    const graph = fakeGraphClient({
      get: async () => {
        calls.push('home');
        return foreignTenantAudienceError;
      },
      discoverTenantId: async (host: string) => {
        calls.push(`discover:${host}`);
        return ok(PARTNER);
      },
      getGuest: async () => {
        calls.push('guest');
        return ok(driveItem);
      },
    });

    const result = await execute(graph, { url: 'https://contoso.sharepoint.com/:p:/s/Team/EaB1cD?e=xyz' });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as Record<string, unknown>;
    expect(value['driveId']).toBe('b!driveId');
    // The whole point: the caller now knows which tenant to name on every
    // follow-up *-drive-item call, which driveId + itemId cannot tell it.
    expect(value['tenantId']).toBe(PARTNER);
    expect(calls).toEqual(['home', 'discover:contoso.sharepoint.com', 'guest']);
  });

  it('reports the tenant boundary plainly when a partner tenant cannot be identified', async () => {
    const graph = fakeGraphClient({
      get: async () => foreignTenantAudienceError,
      discoverTenantId: async () =>
        err({
          type: 'api_error' as const,
          status: 400,
          message: 'could not resolve a tenant for vanity.sharepoint.com (tried vanity.onmicrosoft.com)',
          code: 'tenant_discovery_failed',
        }),
    });

    const result = await execute(graph, { url: 'https://vanity.sharepoint.com/:p:/s/Team/EaB1cD' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type !== 'api_error') return;
    expect(result.error.code).toBe('guest_tenant_unresolved');
    // Repeating the raw `invalidAudienceUri` here would be useless: the reader
    // needs to know it is a tenant boundary AND that the tenant is unidentifiable.
    expect(result.error.message).toContain('guest');
  });

  // A link you simply have no permission to is NOT a tenant problem, and minting a
  // guest token would not help. Retrying it would spend a redemption and turn a
  // clear "no access" into a confusing tenant error.
  it('does not reach for a guest token when the link is merely one you cannot access', async () => {
    let guestCalls = 0;
    const graph = fakeGraphClient({
      get: async () => err({ type: 'api_error' as const, status: 403, message: 'Access denied', code: 'accessDenied' }),
      getGuest: async () => {
        guestCalls += 1;
        return ok(driveItem);
      },
    });

    const result = await execute(graph, { url: 'https://fabrikam.sharepoint.com/:p:/s/Team/EaB1cD' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type !== 'api_error') return;
    expect(result.error.code).toBe('accessDenied');
    expect(guestCalls).toBe(0);
  });

  // The three cases below pin each clause of the "is this the tenant boundary?"
  // guard independently. The accessDenied case above only exercises two of them at
  // once, which lets a mutant delete either one and survive — the same gap the
  // get-user-manager 404 chain carries (LESSONS 2026-07-15). Each is also a real
  // misfire: reaching for a guest token on the wrong error spends a single-use
  // refresh-token redemption and reports a tenant problem that does not exist.
  it('does not reach for a guest token when a 401 is about the token itself rather than the tenant', async () => {
    let guestCalls = 0;
    const graph = fakeGraphClient({
      get: async () => err({ type: 'api_error' as const, status: 401, message: 'Access token has expired.', code: 'InvalidAuthenticationToken' }),
      getGuest: async () => {
        guestCalls += 1;
        return ok(driveItem);
      },
    });

    const result = await execute(graph, { url: 'https://fabrikam.sharepoint.com/:p:/s/Team/EaB1cD' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type !== 'api_error') return;
    expect(result.error.code).toBe('InvalidAuthenticationToken');
    expect(guestCalls).toBe(0);
  });

  it('does not reach for a guest token when an audience complaint arrives with a non-401 status', async () => {
    let guestCalls = 0;
    const graph = fakeGraphClient({
      get: async () => err({ type: 'api_error' as const, status: 403, message: "Invalid audience Uri '00000003-0000-0ff1-ce00-000000000000'.", code: 'invalidAudienceUri' }),
      getGuest: async () => {
        guestCalls += 1;
        return ok(driveItem);
      },
    });

    const result = await execute(graph, { url: 'https://fabrikam.sharepoint.com/:p:/s/Team/EaB1cD' });

    expect(result.ok).toBe(false);
    expect(guestCalls).toBe(0);
  });

  it('does not reach for a guest token when the network failed rather than Graph refusing', async () => {
    let guestCalls = 0;
    const graph = fakeGraphClient({
      get: async () => err({ type: 'network_error' as const, message: 'fetch failed' }),
      getGuest: async () => {
        guestCalls += 1;
        return ok(driveItem);
      },
    });

    const result = await execute(graph, { url: 'https://fabrikam.sharepoint.com/:p:/s/Team/EaB1cD' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('network_error');
    expect(guestCalls).toBe(0);
  });

  it('surfaces the guest attempt failing when the partner tenant refuses the caller', async () => {
    const graph = fakeGraphClient({
      get: async () => foreignTenantAudienceError,
      discoverTenantId: async () => ok(PARTNER),
      getGuest: async () => err({ type: 'auth_failed' as const, message: 'tenant refused a guest token', code: 'secondary_token_unavailable' }),
    });

    const result = await execute(graph, { url: 'https://contoso.sharepoint.com/:p:/s/Team/EaB1cD' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('auth_failed');
  });
});
