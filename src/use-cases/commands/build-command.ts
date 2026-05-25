import type { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import {
  appendOData,
  filterSelectSchema,
  noSkipShape,
  odataQuerySchema,
  pickODataShape,
  selectExpandSchema,
  type FilterSelectParams,
  type ODataKey,
  type ODataQueryParams,
  type SelectExpandParams,
} from './odata-query.ts';

type NoSkipParams = Omit<ODataQueryParams, 'skip'>;

/**
 * Options accepted by every builder that knows about `$select` (i.e. every
 * builder except `buildCommand` / `buildElevatedCommand`, which take no OData
 * passthroughs at all). `defaultSelect`, when set and the user did NOT pass
 * `--select`, is injected into the OData query string so default invocations
 * return a slim projection instead of a 50 KB resource. User-supplied
 * `--select` always wins. Audit Alex-session §A: pairs the `list-mail-attachments`
 * pattern with the builder layer so the 6 heaviest endpoints stop returning
 * the full Graph resource by default.
 */
type SelectDefaults = { readonly defaultSelect?: string };

const withDefaultSelect = <T extends { readonly select?: string }>(data: T, defaultSelect: string | undefined): T => {
  if (defaultSelect === undefined || data.select !== undefined) return data;
  return { ...data, select: defaultSelect };
};

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
  schema: z.ZodObject<Shape>,
  options?: SelectDefaults
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(odataQuerySchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = withDefaultSelect(parsed.data as z.infer<z.ZodObject<Shape>> & ODataQueryParams, options?.defaultSelect);
    const path = appendOData(pathFn(data), data);
    return graph.get(path);
  };
  return { schema: merged, execute };
};

const buildElevatedListCommand = <Shape extends z.ZodRawShape>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>,
  options?: SelectDefaults
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(odataQuerySchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = withDefaultSelect(parsed.data as z.infer<z.ZodObject<Shape>> & ODataQueryParams, options?.defaultSelect);
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
  schema: z.ZodObject<Shape>,
  options?: SelectDefaults
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(selectExpandSchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = withDefaultSelect(parsed.data as z.infer<z.ZodObject<Shape>> & SelectExpandParams, options?.defaultSelect);
    const path = appendOData(pathFn(data), data);
    return graph.get(path);
  };
  return { schema: merged, execute };
};

/**
 * Elevated-token twin of `buildSelectableCommand`. Use for single-resource
 * GETs on endpoints that require the M365ChatClient identity (e.g. `/chats/{id}`)
 * AND benefit from `$select`/`$expand` projection. The basic `buildElevatedCommand`
 * builder takes no OData passthroughs — use this when the endpoint honours
 * field projection, so an LLM can avoid pulling the whole resource just to
 * read a topic or chatType.
 */
const buildElevatedSelectableCommand = <Shape extends z.ZodRawShape>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>,
  options?: SelectDefaults
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(selectExpandSchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = withDefaultSelect(parsed.data as z.infer<z.ZodObject<Shape>> & SelectExpandParams, options?.defaultSelect);
    const path = appendOData(pathFn(data), data);
    return graph.getElevated(path);
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
  schema: z.ZodObject<Shape>,
  options?: SelectDefaults
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(filterSelectSchema.shape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = withDefaultSelect(parsed.data as z.infer<z.ZodObject<Shape>> & FilterSelectParams, options?.defaultSelect);
    const path = appendOData(pathFn(data), data);
    return graph.get(path);
  };
  return { schema: merged, execute };
};

/**
 * Collection GET on an endpoint that supports the usual OData passthroughs
 * EXCEPT `$skip` (e.g. `/me/drive/recent`, `/sites/{id}/lists`,
 * `/me/drive/search`). Graph rejects `$skip` with
 * `invalidRequest: $skip is not supported on this API.`; the CLI mirrors
 * by dropping `--skip` from the advertised flag set.
 */
const buildNoSkipListCommand = <Shape extends z.ZodRawShape>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>,
  options?: SelectDefaults
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(noSkipShape);
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = withDefaultSelect(parsed.data as z.infer<z.ZodObject<Shape>> & NoSkipParams, options?.defaultSelect);
    const path = appendOData(pathFn(data), data);
    return graph.get(path);
  };
  return { schema: merged, execute };
};

/**
 * Collection GET that supports an EXPLICIT subset of OData passthroughs.
 * Use for endpoints where Graph silently drops some flags — passing
 * `keys: ['top', 'select']` advertises only `--top` and `--select` and
 * keeps the manifest honest. The other narrower builders
 * (`buildNoSkipListCommand`, `buildFilterSelectListCommand`) are
 * specializations; this is the generic escape hatch.
 */
const buildPickODataListCommand = <Shape extends z.ZodRawShape, K extends ODataKey>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>,
  keys: ReadonlyArray<K>,
  options?: SelectDefaults
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(pickODataShape(keys));
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = withDefaultSelect(parsed.data as { readonly select?: string } & Record<string, unknown>, options?.defaultSelect);
    const path = appendOData(pathFn(data as z.infer<z.ZodObject<Shape>>), data);
    return graph.get(path);
  };
  return { schema: merged, execute };
};

/**
 * Elevated-token twin of `buildPickODataListCommand`. Use for endpoints that
 * require the M365ChatClient identity (e.g. `/me/chats`, `/chats/{}/members`)
 * AND honour only a subset of OData passthroughs — the chats family rejects
 * `$orderby` / `$expand` with `BadRequest`, so the picker is the right tool.
 */
const buildElevatedPickODataListCommand = <Shape extends z.ZodRawShape, K extends ODataKey>(
  pathFn: (params: z.infer<z.ZodObject<Shape>>) => string,
  schema: z.ZodObject<Shape>,
  keys: ReadonlyArray<K>,
  options?: SelectDefaults
): Pick<Command, 'schema' | 'execute'> => {
  const merged = schema.extend(pickODataShape(keys));
  const execute: Command['execute'] = async (graph, params) => {
    const parsed = merged.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const data = withDefaultSelect(parsed.data as { readonly select?: string } & Record<string, unknown>, options?.defaultSelect);
    const path = appendOData(pathFn(data as z.infer<z.ZodObject<Shape>>), data);
    return graph.getElevated(path);
  };
  return { schema: merged, execute };
};

export {
  buildCommand,
  buildElevatedCommand,
  buildElevatedListCommand,
  buildElevatedPickODataListCommand,
  buildElevatedSelectableCommand,
  buildFilterSelectListCommand,
  buildListCommand,
  buildNoSkipListCommand,
  buildPickODataListCommand,
  buildSelectableCommand,
};
