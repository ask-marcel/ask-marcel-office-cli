import { describe, expect, it } from 'bun:test';
import { TENANT_ID_OPTION } from './tenant-option.ts';

// The `--tenant-id` flag a guest-capable read offers (core, package split phase 3).

describe('the --tenant-id flag', () => {
  it('is optional, keyed tenantId, and says where the tenant id comes from', () => {
    expect({ name: TENANT_ID_OPTION.name, key: TENANT_ID_OPTION.key, required: TENANT_ID_OPTION.required }).toEqual({ name: 'tenant-id', key: 'tenantId', required: false });
    expect(TENANT_ID_OPTION.description).toStartWith('Tenant GUID of a PARTNER tenant you are a guest in');
    expect(TENANT_ID_OPTION.description).toContain('`resolve-drive-share-link`');
  });
});
