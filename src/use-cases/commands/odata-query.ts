import { z } from 'zod';
import type { CommandOptionMeta } from './command-types.ts';

const NON_NEGATIVE_INTEGER = /^\d+$/;

const odataQuerySchema = z.object({
  top: z.string().regex(NON_NEGATIVE_INTEGER, '$top must be a non-negative integer').optional(),
  skip: z.string().regex(NON_NEGATIVE_INTEGER, '$skip must be a non-negative integer').optional(),
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
    description: 'OData $top: maximum number of items to return on this page (non-negative integer). Combine with `next-page` to paginate.',
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

export { appendOData, odataQueryOptions, odataQuerySchema };
export type { ODataQueryParams };
