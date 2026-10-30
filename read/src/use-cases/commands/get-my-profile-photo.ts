import { z } from 'zod';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { ReadGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { inlineBinary } from '@ask-marcel/office-core/use-cases/commands/fetch-raw-bytes.ts';
import { formatZodError } from '@ask-marcel/office-core/use-cases/commands/format-zod-error.ts';

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
  example: '{read-bin} get-my-profile-photo',
  responseShape:
    '`{ contentType: "image/jpeg", size: <bytes>, base64: "<encoded>" }` — the photo bytes, inlined. Pair with the global `--output-path <path>` flag to land the image on disk and replace `base64` with `savedTo`.',
  producesBytes: true,
  effect: 'read',
  scopesRequired: ['User.Read'],
};

export { execute, meta, schema };
