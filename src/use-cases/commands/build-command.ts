import type { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQuerySchema, type ODataQueryParams } from './odata-query.ts';

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

export { buildCommand, buildElevatedCommand, buildElevatedListCommand, buildListCommand };
