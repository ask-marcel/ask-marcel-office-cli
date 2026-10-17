import type { Result } from './result.ts';
import { err, ok } from './result.ts';

export type TenantId = string & { readonly __brand: 'TenantId' };

export type TenantIdError = { type: 'invalid_tenant_id' };

/**
 * An Entra (Azure AD) tenant GUID, used as the authority segment of
 * `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`.
 *
 * That request carries the refresh token in its body, so the segment is a
 * trust boundary (hard rule 12): the host is fixed, but an unvalidated value
 * could still walk the path to a different endpoint on it. Accepting only the
 * canonical 8-4-4-4-12 GUID form closes the category once, here, and every
 * downstream holder of a `TenantId` trusts it.
 *
 * Case is not normalized: Entra accepts either casing on the authority, and
 * the value is also a cache key, so rewriting it would silently split the
 * per-tenant token cache across two entries for one tenant.
 */
const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export const tenantId = (raw: string): Result<TenantId, TenantIdError> => {
  if (!GUID_PATTERN.test(raw)) return err({ type: 'invalid_tenant_id' });
  return ok(raw as TenantId);
};

export const tenantIdUnsafe = (raw: string): TenantId => raw as TenantId;
