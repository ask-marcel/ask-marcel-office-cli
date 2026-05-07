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
    description: 'OData $select: comma-separated list of fields to include in each item (e.g. `id,subject,from`). Shrinks payloads dramatically.',
  },
  {
    name: 'filter',
    key: 'filter',
    required: false,
    description: 'OData $filter: KQL-style predicate to narrow results server-side (e.g. `isRead eq false`). Same syntax Graph documents per resource type.',
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

export { appendOData, filterSelectOptions, filterSelectSchema, odataQueryOptions, odataQuerySchema, selectExpandOptions, selectExpandSchema };
export type { FilterSelectParams, ODataQueryParams, SelectExpandParams };
