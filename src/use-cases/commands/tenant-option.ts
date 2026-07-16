import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import { tenantId } from '../../domain/tenant-id.ts';
import type { TenantId } from '../../domain/tenant-id.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandOptionMeta } from './command-types.ts';

/**
 * `--tenant-id`: read a file that lives in a tenant you are only a GUEST in.
 *
 * Without it, such a read dies at `401 invalidAudienceUri` — home-tenant Graph
 * cannot mint a SharePoint token for a foreign tenant, so no home-tier token can
 * reach it. Unlike `resolve-drive-share-link`, these commands cannot recover on
 * their own: they hold only a `driveId` and an `itemId`, and NEITHER carries a
 * tenant. The sharing URL does, which is why `resolve-drive-share-link` returns
 * the `tenantId` for this flag to consume.
 *
 * OPTIONAL, always. That is not a detail — `required: true` compiles straight to
 * commander's `.requiredOption()`, so an optional flag is what keeps the
 * `CommandMeta` contract untouched and every existing invocation working.
 */
const tenantIdShape = { tenantId: z.string().optional() };

const tenantIdSchema = z.object(tenantIdShape);

const TENANT_ID_OPTION: CommandOptionMeta = {
  name: 'tenant-id',
  key: 'tenantId',
  required: false,
  description:
    'Tenant GUID of a PARTNER tenant you are a guest in, when the file does not live in your own tenant. Get it from `resolve-drive-share-link`, which returns `tenantId` on any sharing URL whose file belongs to another tenant — that is the only place the tenant is knowable, because `--drive-id` and `--item-id` carry no tenant at all. Omit it for files in your own tenant (the normal case). Pass it and the request is signed with a guest token minted for that tenant instead of your home token; without it, a partner-tenant file fails with `invalidAudienceUri` no matter which of your tokens is used.',
};

/**
 * Route a JSON GET to the right identity: the home token normally, a partner
 * tenant's guest token when `--tenant-id` was given.
 *
 * The guest tier is a RUNTIME choice, not a static one like `elevated` — the same
 * command uses either, depending on this flag. That is why the builders route
 * here instead of growing guest twins (which would take 10 builders to 14+, and
 * would still not express "it depends on the argument").
 */
const routeGet = async (graph: GraphClient, path: string, rawTenantId: string | undefined): Promise<Result<unknown, GraphError>> => {
  if (rawTenantId === undefined) return graph.get(path);
  const branded = brandTenantId(rawTenantId);
  if (!branded.ok) return branded;
  return graph.getGuest(path, branded.value);
};

/**
 * Validate a caller-supplied tenant id at the boundary. The value reaches an
 * authority URL whose POST carries the refresh token, so it is branded rather
 * than trusted (hard rule 12); a bad one must fail here with a readable message,
 * not further in.
 */
const brandTenantId = (raw: string): Result<TenantId, GraphError> => {
  const branded = tenantId(raw);
  if (branded.ok) return branded;
  return err({
    type: 'validation_error',
    message: `--tenant-id must be a tenant GUID (8-4-4-4-12), e.g. 6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04. Got: ${raw}. \`resolve-drive-share-link\` returns the right value as \`tenantId\`.`,
    code: 'invalid_tenant_id',
  });
};

export { TENANT_ID_OPTION, brandTenantId, routeGet, tenantIdSchema, tenantIdShape };
