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
    "Decode the cached Teams web client access token and return its scopes, audience, and expiry without making a Graph call. Use this as a self-test before running a command an LLM expects to fail with `accessDenied` — if the required scope isn't in the returned list, the call will reject regardless of tenant config. Each command's `scopesRequired` field in `help-json` lists the scopes that command needs; intersect with the array returned here for a pre-flight check (pipe both through `jq` and diff). The `expiresInSeconds` field lets an LLM decide pre-emptively to `login` again — typically worth doing under ~5 minutes (300 s) so a long-running session doesn't hit the wall mid-command. The `elevated` block reports whether the *separate* M365ChatClient-elevated token (needed by the historical-version download / convert commands) is cached and still usable — so a fresh process can pre-flight `deep-scan`-style workloads instead of discovering a 403 mid-run; `available:false` when it is absent, expired, or within the same 5-minute buffer the download path applies.",
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate: '(meta) cached-token introspection — no Graph endpoint',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/permissions-reference',
  options: [],
  example: 'ask-marcel-office scopes-check',
  responseShape:
    '`{ scopes: string[], audience: string, expiresAt: string (ISO 8601), expiresInSeconds: number, elevated: { available: boolean, expiresInSeconds?: number }, chatsvcagg: { available: boolean, expiresInSeconds?: number }, ic3: { available: boolean, expiresInSeconds?: number } }`. `expiresInSeconds` is negative when the cached token has already expired (run `login`); `audience` is the JWT `aud` claim (typically `https://graph.microsoft.com`). `elevated.available` is `true` only when the cached M365ChatClient-elevated token (used by the historical-version commands) is present and beyond the 5-minute buffer; `elevated.expiresInSeconds` is its raw remaining seconds and is omitted (the key is absent) when no elevated token is cached. `chatsvcagg` and `ic3` are the two Teams-chat substrate tokens, same shape as `elevated`; both self-heal from the shared refresh token, so they are informational rather than a preflight gate.',
};

export { execute, meta, schema };
