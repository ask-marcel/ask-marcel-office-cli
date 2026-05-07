import type { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, filterSelectSchema, odataQuerySchema, selectExpandSchema, type FilterSelectParams, type ODataQueryParams, type SelectExpandParams } from './odata-query.ts';

const buildCommand = (pathFn: (params: Record<string, string>) => string, schema: z.ZodType): Pick<Command, 'schema' | 'execute'> => {
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = schema.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const path = pathFn(parsed.data as Record<string, string>);
    return graph.get(path);
  };
  return { schema, execute };
};

const buildElevatedCommand = (pathFn: (params: Record<string, string>) => string, schema: z.ZodType): Pick<Command, 'schema' | 'execute'> => {
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = schema.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const path = pathFn(parsed.data as Record<string, string>);
    return graph.getElevated(path);
  };
  return { schema, execute };
};

const buildListCommand = <Shape extends z.ZodRawShape>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(odataQuerySchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = parsed.data as z.infer<z.ZodObject<Shape>> & ODataQueryParams;
    const path = appendOData(pathFn(data), data);
    return graph.get(path);
  };
  return { schema: merged, execute };
};

const buildElevatedListCommand = <Shape extends z.ZodRawShape>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(odataQuerySchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = parsed.data as z.infer<z.ZodObject<Shape>> & ODataQueryParams;
    const path = appendOData(pathFn(data), data);
    return graph.getElevated(path);
  };
  return { schema: merged, execute };
};

/**
 * Single-resource GET that supports the OData `$select` and `$expand` query
 * parameters. Mirrors `buildListCommand` but exposes only the two flags that
 * make sense on a non-paginated resource — no `$top`/`$skip`/`$filter`/
 * `$orderby` since there's no collection to slice. Lets an LLM ask only for
 * the fields it needs (e.g. `--select id,subject`) instead of swallowing a
 * 50 KB resource just to read a subject line.
 */
const buildSelectableCommand = <Shape extends z.ZodRawShape>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(selectExpandSchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = parsed.data as z.infer<z.ZodObject<Shape>> & SelectExpandParams;
    const path = appendOData(pathFn(data), data);
    return graph.get(path);
  };
  return { schema: merged, execute };
};

/**
 * Collection GET that supports ONLY `$filter` and `$select` — for endpoints
 * Microsoft documents as rejecting the other OData passthroughs (`/teams/{id}/channels`
 * is the canonical case: Graph returns BadRequest on `$top`, `$skip`, `$orderby`,
 * `$expand`). Advertising the unsupported flags would be a usability lie.
 */
const buildFilterSelectListCommand = <Shape extends z.ZodRawShape>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(filterSelectSchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = parsed.data as z.infer<z.ZodObject<Shape>> & FilterSelectParams;
    const path = appendOData(pathFn(data), data);
    return graph.get(path);
  };
  return { schema: merged, execute };
};

export { buildCommand, buildElevatedCommand, buildElevatedListCommand, buildFilterSelectListCommand, buildListCommand, buildSelectableCommand };
