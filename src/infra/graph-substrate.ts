import type { AccessToken } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn, GraphError } from './graph-request.ts';
import { apiErrorFrom, asAuthFailure, toBase64, wrapNetworkError } from './graph-request.ts';
import { REQUEST_TIMEOUT_MS, timeoutMsFor } from './network-error.ts';

/*
 * The Teams chat substrate reads of the read graph (read-graph.ts): chat
 * bodies on chatsvcagg and IC3, and a pasted chat image from the Teams media
 * service. Each reads only.
 */

type SubstrateReader = {
  readonly teamsChat: (path: string) => Promise<Result<unknown, GraphError>>;
  readonly teamsChatIc3: (path: string) => Promise<Result<unknown, GraphError>>;
  readonly teamsChatMedia: (url: string) => Promise<Result<unknown, GraphError>>;
};

// The Teams media service hosts that may receive the IC3 bearer; a chat image
// URL is read out of message content, so nothing else is trusted with it.
const TEAMS_MEDIA_HOST = /^[a-z0-9-]+\.(?:asm\.skype\.com|asyncgw\.teams\.microsoft\.com)$/i;

const teamsMediaHost = (url: string): string | undefined => {
  try {
    const parsed = new URL(url);
    return parsed.protocol === 'https:' && TEAMS_MEDIA_HOST.test(parsed.hostname) ? parsed.hostname : undefined;
  } catch {
    return undefined;
  }
};

/**
 * Tag an api_error returned from a Microsoft-internal chat substrate
 * (chatsvcagg `/api/csa/<region>/...` or IC3 `/api/chatsvc/<region>/...`)
 * with a `substrateHttp{status}_{substrate}` code so the presenter's hint
 * table can match it and add the "best-effort substrate may have moved"
 * actionable hint plus `source: "substrate"` classifier. Prior shape left
 * substrate errors with whatever code (or
 * none) Graph returned, indistinguishable from regular Graph errors and
 * without the experimental-substrate context an LLM needs to decide whether
 * to retry, switch substrates, or surface the failure to the user.
 *
 * Non-api_error inputs (auth_failed, validation_error, network errors) pass
 * through unchanged — those are upstream-of-substrate failures and the
 * existing classifier handles them.
 */
const asSubstrateError = (e: GraphError, substrate: 'chatsvcagg' | 'ic3'): GraphError => {
  if (e.type !== 'api_error') return e;
  return { ...e, code: `substrateHttp${e.status}_${substrate}` };
};

const createSubstrateReader = (tokens: TokenSource, fetchFn: FetchFn): SubstrateReader => {
  // Teams chat substrate. Same Teams web client identity as `get`, but the
  // bearer is issued for `chatsvcagg.teams.microsoft.com` (audience claim only
  // — the actual API lives on `teams.microsoft.com/api/csa/<region>/` since the
  // 2026-05 substrate move) or for `https://ic3.teams.office.com` on the IC3
  // path prefix. We piggy-back the captured bearer to read chat message bodies:
  // Graph's `Chat.Read*`-gated endpoints cannot reach them with the scopes the
  // basic Teams token carries. IC3 is the one Teams web actually uses for
  // scrollback, since it supports `syncState` + `startTime` pagination that
  // chatsvcagg lacks. See `gotcha_chatsvcagg_substrate_moved` in memory.
  //
  // A 401 here is read as "this token is DEAD", not as an answer. A substrate
  // token can be revoked server-side while still inside its expiry window — a
  // second sign-in invalidates the previous session's — and nothing in the
  // cache records that, so the tier keeps reporting available while every chat
  // command 401s, and `login` cannot recover it because the token is not
  // missing (observed live 2026-08-31: a token minted at 07:18 was rejected at
  // 13:47 with hours of stated life left, while a freshly redeemed one worked
  // instantly). So: drop it, redeem a fresh one from the shared refresh token
  // over HTTP, replay ONCE. A second 401 is real and is surfaced. Any other
  // status returns as-is, because no amount of fresh token fixes a 404.
  // One request signed with a substrate bearer. A 401 sends it once more with a
  // token the source mints past the rejected one; any other answer stands.
  const sendWithSubstrateToken = async (kind: 'chatsvcagg' | 'ic3', send: (authorization: string) => Promise<Response>): Promise<Result<Response, GraphError>> => {
    const attempt = async (rejected?: AccessToken): Promise<Result<{ readonly response: Response; readonly token: AccessToken }, GraphError>> => {
      const token = await tokens.substrateToken(kind, rejected === undefined ? undefined : { rejected });
      if (!token.ok) return err(asAuthFailure(token.error));
      return ok({ response: await send(`Bearer ${token.value}`), token: token.value });
    };
    const first = await attempt();
    if (!first.ok) return first;
    if (first.value.response.status !== 401) return ok(first.value.response);
    const replay = await attempt(first.value.token);
    return replay.ok ? ok(replay.value.response) : replay;
  };

  const substrateGet = async (kind: 'chatsvcagg' | 'ic3', prefix: 'csa' | 'chatsvc', path: string): Promise<Result<unknown, GraphError>> => {
    const region = await tokens.substrateRegion(kind);
    if (!region.ok) return err(asAuthFailure(region.error));
    const url = `https://teams.microsoft.com/api/${prefix}/${region.value}${path}`;
    try {
      const sent = await sendWithSubstrateToken(kind, (authorization) =>
        fetchFn(url, { method: 'GET', headers: { Authorization: authorization, accept: 'application/json' }, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) })
      );
      if (!sent.ok) return sent;
      if (!sent.value.ok) return err(asSubstrateError(await apiErrorFrom(sent.value, url), kind));
      return ok(await sent.value.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${path} (${kind})`, 'json'));
    }
  };

  const teamsChatMedia = async (url: string): Promise<Result<unknown, GraphError>> => {
    const host = teamsMediaHost(url);
    if (host === undefined)
      return err({ type: 'validation_error', message: `not a Teams media URL (https on *.asm.skype.com or *.asyncgw.teams.microsoft.com): ${url.slice(0, 120)}` });
    try {
      // A revoked substrate token still looks valid in the cache: a 401 gets one
      // replay with a freshly redeemed token, as on the other substrate reads.
      const sent = await sendWithSubstrateToken('ic3', (authorization) =>
        fetchFn(url, { method: 'GET', headers: { Authorization: authorization }, signal: AbortSignal.timeout(timeoutMsFor('binary')) })
      );
      if (!sent.ok) return sent;
      if (!sent.value.ok) return err(await apiErrorFrom(sent.value, url));
      const buffer = await sent.value.arrayBuffer();
      return ok({ contentType: sent.value.headers.get('content-type') ?? 'application/octet-stream', size: buffer.byteLength, base64: toBase64(new Uint8Array(buffer)) });
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${host} (teams media)`, 'binary'));
    }
  };

  return {
    teamsChat: (path) => substrateGet('chatsvcagg', 'csa', path),
    teamsChatIc3: (path) => substrateGet('ic3', 'chatsvc', path),
    teamsChatMedia,
  };
};

export { createSubstrateReader };
export type { SubstrateReader };
