import { z } from 'zod';
import type { CommandOptionMeta } from './command-types.ts';

const POSITIVE_INTEGER = /^[1-9]\d*$/;
const NON_NEGATIVE_INTEGER = /^\d+$/;
const SIGNED_INTEGER = /^-?\d+$/;

// Graph silently caps `$top` at 1000 across nearly every collection endpoint —
// a request like `?$top=999999` returns the first 1000 items with NO warning,
// which is a usability trap for LLMs trying to "get everything in one call".
// Reject any `--top` over the cap at the schema level so the operator gets a
// clear error instead of a silent truncation, and document the cap in the
// validation message.
const TOP_HARD_CAP = 1000;

const boundedPositiveIntegerSchema = (label: string, max: number): z.ZodString =>
  z.string().superRefine((value, ctx) => {
    if (!POSITIVE_INTEGER.test(value)) {
      const reason = SIGNED_INTEGER.test(value)
        ? `must be a positive integer (Graph rejects ${label}=0 and negatives)`
        : `must be a positive integer (got "${value}", which is not a number)`;
      ctx.addIssue({ code: 'custom', message: reason });
      return;
    }
    if (Number.parseInt(value, 10) > max) {
      ctx.addIssue({
        code: 'custom',
        message: `must be ≤ ${max} (Graph silently caps ${label} at ${max} on every collection endpoint; pass a smaller value or paginate via next-page)`,
      });
    }
  }) as unknown as z.ZodString;

const nonNegativeIntegerSchema = (label: string): z.ZodString =>
  z.string().superRefine((value, ctx) => {
    if (NON_NEGATIVE_INTEGER.test(value)) return;
    const reason = SIGNED_INTEGER.test(value)
      ? `must be a non-negative integer (Graph rejects ${label}=negatives)`
      : `must be a non-negative integer (got "${value}", which is not a number)`;
    ctx.addIssue({ code: 'custom', message: reason });
  }) as unknown as z.ZodString;

const odataQuerySchema = z.object({
  top: boundedPositiveIntegerSchema('top', TOP_HARD_CAP).optional(),
  skip: nonNegativeIntegerSchema('skip').optional(),
  select: z.string().min(1).optional(),
  filter: z.string().min(1).optional(),
  orderby: z.string().min(1).optional(),
  expand: z.string().min(1).optional(),
});

type ODataQueryParams = z.infer<typeof odataQuerySchema>;

const ORDERED_KEYS = ['top', 'skip', 'select', 'filter', 'orderby', 'expand'] as const;

const appendOData = (path: string, params: ODataQueryParams): string => {
  const parts: string[] = [];
  for (const key of ORDERED_KEYS) {
    const value = params[key];
    if (value === undefined) continue;
    parts.push(`$${key}=${encodeURIComponent(value)}`);
  }
  if (parts.length === 0) return path;
  const separator = path.includes('?') ? '&' : '?';
  return `${path}${separator}${parts.join('&')}`;
};

const odataQueryOptions: ReadonlyArray<CommandOptionMeta> = [
  {
    name: 'top',
    key: 'top',
    required: false,
    description:
      'OData $top: maximum number of items to return on this page (positive integer, ≤ 1000). Graph silently caps at 1000 on every collection endpoint, so the CLI rejects larger values with a clear validation error rather than letting the request silently truncate. Combine with `next-page` to paginate beyond the cap.',
  },
  {
    name: 'skip',
    key: 'skip',
    required: false,
    description: 'OData $skip: skip the first N items before returning results (non-negative integer). Useful with $top for offset paging.',
  },
  {
    name: 'select',
    key: 'select',
    required: false,
    description:
      'OData $select: comma-separated list of fields to include in each item (e.g. `id,subject,from`). May shrink payloads dramatically — Graph honors $select on most endpoints, but some collections (notably `/me/mailboxSettings`, `/me/outlook/masterCategories`, `/me/mailFolders/inbox/messageRules`) silently ignore it and always return the full resource.',
  },
  {
    name: 'filter',
    key: 'filter',
    required: false,
    description:
      "OData $filter: predicate to narrow results server-side. Quoting rules: string literals MUST use SINGLE quotes (`subject eq 'invoice'`), NOT double quotes — Graph rejects `subject eq \"invoice\"` with `InvalidFilterClause`. To embed a single quote inside a string, double it (`subject eq 'O''Brien'`). Booleans, numbers, and dates are unquoted (`isRead eq false`, `receivedDateTime ge 2026-01-01T00:00:00Z`). Wrap the whole flag value in shell DOUBLE quotes so the inner single quotes survive (`--filter \"subject eq 'invoice'\"`). Same syntax Graph documents per resource type.",
  },
  {
    name: 'orderby',
    key: 'orderby',
    required: false,
    description:
      'OData $orderby: sort expression with optional asc/desc (e.g. `receivedDateTime desc`). Some Graph filter combinations are rejected; remove $orderby if InefficientFilter occurs.',
  },
  {
    name: 'expand',
    key: 'expand',
    required: false,
    description: 'OData $expand: navigation properties to include inline (e.g. `attachments`). Increases response size; use sparingly.',
  },
];

/**
 * Subset for single-resource GETs that support `$select` and `$expand` but
 * have no collection to slice (`$top`/`$skip`/`$filter`/`$orderby` make no
 * sense). Keeping this as a separate Zod fragment lets `buildSelectableCommand`
 * advertise only the two relevant flags in `--help`.
 */
const selectExpandSchema = z.object({
  select: z.string().min(1).optional(),
  expand: z.string().min(1).optional(),
});

type SelectExpandParams = z.infer<typeof selectExpandSchema>;

const selectExpandOptions: ReadonlyArray<CommandOptionMeta> = odataQueryOptions.filter((o) => o.name === 'select' || o.name === 'expand');

/**
 * Subset for collection endpoints that Microsoft documents as supporting only
 * `$filter` + `$select` (e.g. `/teams/{id}/channels`). Advertising the other
 * OData passthroughs would be a usability lie since Graph rejects them with
 * `BadRequest`.
 */
const filterSelectSchema = z.object({
  filter: z.string().min(1).optional(),
  select: z.string().min(1).optional(),
});

type FilterSelectParams = z.infer<typeof filterSelectSchema>;

const filterSelectOptions: ReadonlyArray<CommandOptionMeta> = odataQueryOptions.filter((o) => o.name === 'filter' || o.name === 'select');

/**
 * Subset for collection endpoints that REJECT `$skip` with
 * `invalidRequest: $skip is not supported on this API. Only URLs returned
 * by the API can be used to page.` (e.g. `/sites/{id}/lists`). Drop --skip
 * from the advertised flags; pagination still works via `nextLink` →
 * `next-page`.
 */
const noSkipShape = Object.fromEntries(Object.entries(odataQuerySchema.shape).filter(([key]) => key !== 'skip')) as Omit<typeof odataQuerySchema.shape, 'skip'>;

const noSkipOptions: ReadonlyArray<CommandOptionMeta> = odataQueryOptions.filter((o) => o.name !== 'skip');

type ODataKey = keyof typeof odataQuerySchema.shape;

/**
 * Returns a zod-extension shape containing ONLY the named OData fields.
 * Use for endpoints where Graph silently ignores some passthroughs — the
 * CLI advertises only the ones the endpoint honors.
 */
const pickODataShape = <K extends ODataKey>(keys: ReadonlyArray<K>): Pick<typeof odataQuerySchema.shape, K> => {
  const out: Partial<typeof odataQuerySchema.shape> = {};
  for (const key of keys) out[key] = odataQuerySchema.shape[key];
  return out as Pick<typeof odataQuerySchema.shape, K>;
};

/**
 * Returns the matching CommandOptionMeta entries for use in `meta.options`.
 * Order is preserved from the canonical `odataQueryOptions` definition.
 */
const pickODataOptions = (keys: ReadonlyArray<ODataKey>): ReadonlyArray<CommandOptionMeta> => {
  const allowed = new Set<string>(keys);
  return odataQueryOptions.filter((o) => allowed.has(o.name));
};

/**
 * Subset for delta-tracking endpoints where Graph silently ignores every
 * OData passthrough except `$top` (which itself must be translated to a
 * `Prefer: odata.maxpagesize` header — `$top` as a query parameter returns
 * `ErrorInvalidUrlQuery`). The CLI keeps `--top` as the user-facing flag
 * and drops the others rather than advertising no-ops.
 */
const topOnlyShape = pickODataShape(['top']);
const topOnlyOptions = pickODataOptions(['top']);

/**
 * Subset for endpoints where Graph honors ONLY `$select` (e.g.
 * `/me/planner/plans`, `/planner/plans/{id}/buckets` — the other OData
 * passthroughs are silently dropped server-side).
 */
const selectOnlyShape = pickODataShape(['select']);
const selectOnlyOptions = pickODataOptions(['select']);

export {
  appendOData,
  filterSelectOptions,
  filterSelectSchema,
  noSkipOptions,
  noSkipShape,
  odataQueryOptions,
  odataQuerySchema,
  pickODataOptions,
  pickODataShape,
  selectExpandOptions,
  selectExpandSchema,
  selectOnlyOptions,
  selectOnlyShape,
  topOnlyOptions,
  topOnlyShape,
};
export type { FilterSelectParams, ODataKey, ODataQueryParams, SelectExpandParams };
