import { describe, expect, it } from 'bun:test';
import { tenantId, tenantIdUnsafe } from './tenant-id.ts';

describe('tenantId brand factory', () => {
  it('accepts a canonical tenant GUID as an authority segment', () => {
    const result = tenantId('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04'));
  });

  it('accepts a tenant GUID regardless of the casing Microsoft returns it in', () => {
    const result = tenantId('6F1E3A92-4B7C-4D51-9E2F-8A3B5C7D1E04');
    expect(result.ok).toBe(true);
  });

  // The brand exists for this case. The value is interpolated into
  // `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`, and that
  // request carries the refresh token in its body. The host is fixed, so this is
  // not exfiltration, but a traversal segment would redirect the POST to another
  // endpoint on the same host. The GUID gate closes the category at the boundary.
  it('refuses a tenant id carrying a path traversal before it can redirect the refresh-token POST', () => {
    const result = tenantId('../../evil/oauth2/v2.0/token');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('invalid_tenant_id');
  });

  // The dangerous shape is a traversal APPENDED to a real GUID: the case above
  // carries no GUID at all, so the character classes reject it whatever the
  // anchors do — it passes for the wrong reason. Only this one pins the trailing
  // `$`, and without that anchor the POST that carries the refresh token walks to
  // an attacker-chosen path. (Both anchors surfaced as surviving Regex mutants.)
  it('refuses a path traversal appended to an otherwise valid tenant GUID', () => {
    const result = tenantId('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04/../../evil');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('invalid_tenant_id');
  });

  it('refuses a tenant GUID that has been prefixed with another path segment', () => {
    const result = tenantId('evil/6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('invalid_tenant_id');
  });

  it('refuses a tenant domain, because only the GUID form addresses the authority', () => {
    const result = tenantId('contoso.onmicrosoft.com');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('invalid_tenant_id');
  });

  it('refuses an empty tenant id', () => {
    const result = tenantId('');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('invalid_tenant_id');
  });
});
