import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({}).strict();

// A single, jargon-free action line: `login --force` re-captures every token.
const REFRESH_HINT = 'To refresh any token, run `ask-marcel-office login --force`.';

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const info = await graph.getCachedTokenInfo();
  if (!info.ok) return info;
  return ok({ ...info.value, hint: REFRESH_HINT });
};

const meta: CommandMeta = {
  summary:
    "Decode the cached Teams web client access token and return its scopes, audience, and expiry without making a Graph call. Use this as a self-test before running a command an LLM expects to fail with `accessDenied` — if the required scope isn't in the returned list, the call will reject regardless of tenant config. Each command's `scopesRequired` field in `help-json` lists the scopes that command needs; intersect with the array returned here for a pre-flight check (pipe both through `jq` and diff). The `expiresInSeconds` field lets an LLM decide pre-emptively to `login` again — typically worth doing under ~5 minutes (300 s) so a long-running session doesn't hit the wall mid-command. The `elevated` block reports whether the *separate* M365ChatClient-elevated token (needed by the historical-version download / convert commands) is cached and still usable — so a fresh process can pre-flight `deep-scan`-style workloads instead of discovering a 403 mid-run; `available:false` when it is absent, expired, or within the same 5-minute buffer the download path applies. The `chatsvcagg` and `ic3` blocks report the two Teams-chat substrate tokens (used by `list-teams-chat*` / `find-chats-with-user`) the same way; both self-heal from the shared refresh token, so they are informational rather than a preflight gate. Every tier block also lists that token's OWN granted scopes (decoded from its `scp` claim, distinct per token) and its `refresh` route (`automatic` vs `interactive`); the `hint` field names a forced re-login as the single way to refresh any of them.",
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate: '(meta) cached-token introspection — no Graph endpoint',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/permissions-reference',
  options: [],
  example: 'ask-marcel-office scopes-check',
  responseShape:
    '`{ scopes: string[], audience: string, expiresAt: string (ISO 8601), expiresInSeconds: number, elevated: TokenTier, chatsvcagg: TokenTier, ic3: TokenTier, hint: string }` where `TokenTier = { available: boolean, expiresInSeconds?: number, scopes: string[], refresh: "automatic" | "interactive", reason?: string }`. Top-level `scopes`/`audience`/`expiresAt`/`expiresInSeconds` describe the basic Teams token (back-compat). Each tier block also carries that token\'s OWN `scopes` (decoded from its `scp` claim: elevated ~20 Graph scopes, chatsvcagg `user_impersonation`, ic3 `Teams.AccessAsUser.All`) and its `refresh` route (`automatic` = self-heals from the shared refresh token; `interactive` = the elevated token, re-captured only by a browser login). `available` is `true` only when the token is present and beyond the 5-minute buffer; `expiresInSeconds` is the raw remaining seconds (negative when expired) and is omitted when the token is absent. `reason` is present ONLY when `available` is `false` — a plain-language note on why the tier is missing and how to restore it (so an empty `scopes: []` on an absent token is not mistaken for "no scopes"); it is omitted when the token is available. `hint` names `login --force` as the single refresh action.',
};

export { execute, meta, schema };
