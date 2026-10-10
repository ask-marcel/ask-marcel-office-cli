import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { ReadGraph } from '../../infra/read-graph.ts';
import type { ReadCommandMeta } from './command-types.ts';
import { inlineBinary } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({}).strict();

const execute = async (graph: ReadGraph, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return inlineBinary(graph, '/me/photo/$value');
};

const meta: ReadCommandMeta = {
  summary:
    "Download the signed-in user's profile photo (largest available size), inlined. The CLI follows the Graph 302 → CDN redirect internally so the LLM never has to fetch an external URL.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/photo/$value',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/profilephoto-get',
  options: [],
  example: 'ask-marcel-office get-my-profile-photo',
  responseShape:
    '`{ contentType: "image/jpeg", size: <bytes>, base64: "<encoded>" }` — the photo bytes, inlined. Pair with the global `--output-path <path>` flag to land the image on disk and replace `base64` with `savedTo`.',
  producesBytes: true,
  effect: 'read',
  scopesRequired: ['User.Read'],
};

export { execute, meta, schema };
