import { z } from 'zod';

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

export { appendOData, odataQuerySchema };
export type { ODataQueryParams };
