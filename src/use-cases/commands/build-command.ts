import type { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

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

export { buildCommand, buildElevatedCommand };
