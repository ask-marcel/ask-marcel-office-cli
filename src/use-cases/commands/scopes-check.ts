import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({}).strict();

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return graph.getCachedTokenInfo();
};

const meta: CommandMeta = {
  summary:
    "Decode the cached Teams web client access token and return its scopes, audience, and expiry without making a Graph call. Use this as a self-test before running a command an LLM expects to fail with `accessDenied` — if the required scope isn't in the returned list, the call will reject regardless of tenant config. Each command's `scopesRequired` field in `help-json` lists the scopes that command needs; intersect with the array returned here for a pre-flight check (pipe both through `jq` and diff). The `expiresInSeconds` field (added Alex-session §4) lets an LLM decide pre-emptively to `login` again — typically worth doing under ~5 minutes (300 s) so a long-running session doesn't hit the wall mid-command.",
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate: '(meta) cached-token introspection — no Graph endpoint',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/permissions-reference',
  options: [],
  example: 'ask-marcel scopes-check',
  responseShape:
    '`{ scopes: string[], audience: string, expiresAt: string (ISO 8601), expiresInSeconds: number }`. `expiresInSeconds` is negative when the cached token has already expired (run `login`); `audience` is the JWT `aud` claim (typically `https://graph.microsoft.com`).',
};

export { execute, meta, schema };
