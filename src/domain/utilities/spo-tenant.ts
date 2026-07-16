/**
 * Maps a SharePoint / OneDrive host to the tenant domain that owns it, so a
 * sharing URL alone is enough to discover which tenant to ask for a guest token.
 *
 * A tenant's SharePoint hostname is derived from its initial `*.onmicrosoft.com`
 * domain: `contoso.onmicrosoft.com` serves `contoso.sharepoint.com` for sites and
 * `contoso-my.sharepoint.com` for personal OneDrive. Reversing that gives a
 * domain the Entra OIDC discovery endpoint resolves to a tenant id, with no
 * credentials and no prior knowledge of the tenant.
 *
 * Returns null for anything that is not a tenant SharePoint host — including
 * `1drv.ms`, Microsoft's CONSUMER short-link host, which belongs to no Entra
 * tenant. A null answer means "no guest token applies here", which leaves the
 * caller on its home token: the current behaviour, not a failure.
 *
 * Verified against a real partner tenant 2026-07-16. This mapping is a
 * convention rather than a guarantee, so callers must treat a failed discovery
 * as "cannot cross to this tenant" and surface it, never as a crash.
 */
export const spoHostToTenantDomain = (host: string): string | null => {
  const match = /^([a-z0-9][a-z0-9-]*?)(-my|-admin)?\.sharepoint\.com$/i.exec(host);
  if (match === null) return null;
  const prefix = match[1];
  if (!prefix) return null;
  return `${prefix.toLowerCase()}.onmicrosoft.com`;
};
