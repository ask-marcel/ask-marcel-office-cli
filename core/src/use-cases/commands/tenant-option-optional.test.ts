import { describe, expect, it } from 'bun:test';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import { brandOptionalTenantId } from './tenant-option.ts';

describe('the optional --tenant-id of a command', () => {
  it('is no tenant when the flag is absent', () => {
    expect(brandOptionalTenantId(undefined)).toEqual({ ok: true, value: undefined });
  });

  it('is the branded tenant when the flag holds a GUID', () => {
    expect(brandOptionalTenantId('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04')).toEqual({ ok: true, value: tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04') });
  });

  it('is the invalid_tenant_id error when the flag holds anything else', () => {
    const result = brandOptionalTenantId('contoso');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toEqual(expect.objectContaining({ type: 'validation_error', code: 'invalid_tenant_id' }));
  });
});
