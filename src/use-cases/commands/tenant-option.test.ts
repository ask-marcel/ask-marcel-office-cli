import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '../../domain/result.ts';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildCommand, buildElevatedListCommand, buildListCommand } from './build-command.ts';
import { tenantIdShape } from './tenant-option.ts';

const PARTNER = tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04');

describe('reading a file that lives in a partner tenant', () => {
  it('signs the request with the home token when no partner tenant is named', async () => {
    const seen: string[] = [];
    const graph = fakeGraphClient({
      get: async () => {
        seen.push('home');
        return ok({ id: 'x' });
      },
      getGuest: async () => {
        seen.push('guest');
        return ok({ id: 'x' });
      },
    });
    const { execute } = buildCommand((p) => `/drives/${p['driveId']}/items/${p['itemId']}`, z.object({ driveId: z.string(), itemId: z.string(), tenantId: z.string().optional() }));

    await execute(graph, { driveId: 'b!x', itemId: '01A' });

    expect(seen).toEqual(['home']);
  });

  it('signs the request with a guest token for the named partner tenant', async () => {
    let guestTenant: string | undefined;
    const graph = fakeGraphClient({
      getGuest: async (_path: string, tenant) => {
        guestTenant = tenant;
        return ok({ id: 'x' });
      },
    });
    const { execute } = buildCommand((p) => `/drives/${p['driveId']}/items/${p['itemId']}`, z.object({ driveId: z.string(), itemId: z.string(), tenantId: z.string().optional() }));

    const result = await execute(graph, { driveId: 'b!x', itemId: '01A', tenantId: PARTNER });

    expect(result.ok).toBe(true);
    expect(guestTenant).toBe(PARTNER);
  });

  // The value arrives from the command line and ends up as the authority segment
  // of a URL whose POST body carries the refresh token, so it is branded at this
  // boundary (hard rule 12) rather than trusted. A bad one must fail HERE with a
  // message naming where the right value comes from — not deeper, as a mystery.
  it('refuses a tenant id that is not a tenant GUID, pointing at where to get one', async () => {
    let guestCalls = 0;
    const graph = fakeGraphClient({
      getGuest: async () => {
        guestCalls += 1;
        return ok({});
      },
    });
    const { execute } = buildCommand((p) => `/drives/${p['driveId']}/items/${p['itemId']}`, z.object({ driveId: z.string(), itemId: z.string(), tenantId: z.string().optional() }));

    const result = await execute(graph, { driveId: 'b!x', itemId: '01A', tenantId: '../../evil' });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    if (result.error.type !== 'validation_error') return;
    expect(result.error.code).toBe('invalid_tenant_id');
    expect(result.error.message).toContain('resolve-drive-share-link');
    expect(guestCalls).toBe(0);
  });

  it('carries a partner tenant through a list command alongside its OData flags', async () => {
    let guestPath: string | undefined;
    const graph = fakeGraphClient({
      getGuest: async (path: string) => {
        guestPath = path;
        return ok({ value: [] });
      },
    });
    // A command OPTS IN by declaring the field itself. The builders route, they do
    // not inject: the meta invariant requires one option per schema field, so a
    // builder that added `tenantId` to every schema would force `--tenant-id` onto
    // ~150 commands — including mail and calendar, where a guest token is
    // meaningless. Opting in per command keeps the flag where it can do work.
    const { execute } = buildListCommand((p) => `/drives/${p.driveId}/items`, z.object({ driveId: z.string(), ...tenantIdShape }));

    await execute(graph, { driveId: 'b!x', tenantId: PARTNER, top: '5' });

    expect(guestPath).toBe('/drives/b!x/items?$top=5');
  });

  it('ignores a partner tenant on a command that has not opted in, rather than silently signing as a guest', async () => {
    const seen: string[] = [];
    const graph = fakeGraphClient({
      get: async () => {
        seen.push('home');
        return ok({ value: [] });
      },
      getGuest: async () => {
        seen.push('guest');
        return ok({ value: [] });
      },
    });
    const { execute } = buildListCommand((p) => `/me/messages/${p.id}`, z.object({ id: z.string() }));

    await execute(graph, { id: 'm1', tenantId: PARTNER });

    // Unreachable in practice — commander rejects a flag the command never
    // declared — but the schema must not quietly honour it either.
    expect(seen).toEqual(['home']);
  });

  // The elevated tier is the home-tenant M365ChatClient identity, used for
  // ODSP-gated endpoints. "Elevated in a partner tenant" is not a thing, so the
  // elevated builders never route to guest — the combination is unrepresentable
  // rather than a confusing wire failure.
  it('never reaches for a guest token from an elevated command, even if a tenant is named', async () => {
    const seen: string[] = [];
    const graph = fakeGraphClient({
      getElevated: async () => {
        seen.push('elevated');
        return ok({ value: [] });
      },
      getGuest: async () => {
        seen.push('guest');
        return ok({ value: [] });
      },
    });
    const { execute } = buildElevatedListCommand((p) => `/chats/${p.chatId}/members`, z.object({ chatId: z.string() }));

    await execute(graph, { chatId: '19:abc', tenantId: PARTNER });

    expect(seen).toEqual(['elevated']);
  });
});
