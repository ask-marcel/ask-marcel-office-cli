import type { AccessToken } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import { teamsRegion } from '../domain/teams-region.ts';
import { tokenFingerprint } from '../domain/token-fingerprint.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';
import type { HelperToken } from './token-helper-answer.ts';
import { tierCode } from './token-helper-answer.ts';
import type { TokenHelperRequest } from './token-helper-run.ts';

/*
 * Tokens from the token helper (src/infra/token-helper-run.ts). One run at a
 * time per tier and tenant, shared by every caller that asks meanwhile. An
 * answer is kept in memory for five minutes at most, and never once the helper
 * itself would stop handing it out, which bounds how long a long-running server
 * keeps a bearer after a logout or an account switch. Failures are not kept.
 */

export type HelperTokenSourceDeps = {
  // One run of the helper.
  readonly ask: (request: TokenHelperRequest) => Promise<Result<HelperToken, TokenError>>;
  readonly now?: () => number;
};

type Answer = Result<HelperToken, TokenError>;
type Held = HelperToken & { readonly until: number };

const REUSE_MS = 5 * 60_000;
// The helper's ladder stops handing out a token this close to its expiry.
const EXPIRY_MARGIN_MS = 5 * 60_000;

const keyOf = (request: TokenHelperRequest): string => (request.tenant === undefined ? request.tier : `${request.tier} ${request.tenant}`);

export const createHelperTokenSource = (deps: HelperTokenSourceDeps): TokenSource => {
  const now = deps.now ?? Date.now;
  const held = new Map<string, Held>();
  const inFlight = new Map<string, Promise<Answer>>();
  // Counts the replays per key. A run that a replay started after does not
  // write its answer to memory: it may be the token the replay was refused.
  const replays = new Map<string, number>();

  const keep = (key: string, answer: HelperToken): void => {
    held.set(key, { ...answer, until: Math.min(now() + REUSE_MS, answer.expiresOn * 1000 - EXPIRY_MARGIN_MS) });
  };

  const settle = (key: string, flight: string, generation: number | undefined, answer: Answer): Answer => {
    inFlight.delete(flight);
    if (answer.ok && replays.get(key) === generation) keep(key, answer.value);
    return answer;
  };

  // Callers that ask while a run is under way share it; nothing awaits between
  // the look-up and the entry, so a burst of callers sees one run. A replay is
  // its own run, never joined to a plain one that may hand back the refused token.
  const shared = (request: TokenHelperRequest): Promise<Answer> => {
    const key = keyOf(request);
    const flight = request.rejected === undefined ? key : `${key} ${request.rejected}`;
    const running = inFlight.get(flight);
    if (running !== undefined) return running;
    if (request.rejected !== undefined) replays.set(key, (replays.get(key) ?? 0) + 1);
    const generation = replays.get(key);
    const started = deps.ask(request).then((answer) => settle(key, flight, generation, answer));
    inFlight.set(flight, started);
    return started;
  };

  // A refused token is never handed back: a newer one in memory answers, and
  // otherwise the helper is asked past the refused one by its fingerprint.
  // Memory is read after the fingerprint is taken, so a replay that finished
  // meanwhile answers the callers refused with the same token.
  const answerFor = async (request: TokenHelperRequest, rejected?: AccessToken): Promise<Answer> => {
    const fingerprint = rejected === undefined ? undefined : await tokenFingerprint(rejected);
    const kept = held.get(keyOf(request));
    if (kept !== undefined && now() < kept.until && kept.token !== rejected) return ok(kept);
    return shared(fingerprint === undefined ? request : { ...request, rejected: fingerprint });
  };

  const tokenFor = async (request: TokenHelperRequest, rejected?: AccessToken): Promise<Result<AccessToken, TokenError>> => {
    const answer = await answerFor(request, rejected);
    return answer.ok ? ok(answer.value.token) : answer;
  };

  return {
    graphToken: (tier, options) => tokenFor({ tier }, options?.rejected),
    guestToken: (tenant, options) => tokenFor({ tier: 'guest', tenant }, options?.rejected),
    substrateToken: (tier, options) => tokenFor({ tier }, options?.rejected),
    // Checked here, not when the token is read: the Teams media service takes
    // the ic3 token with no region, and only a URL carries one.
    substrateRegion: async (tier) => {
      const answer = await answerFor({ tier });
      if (!answer.ok) return answer;
      const region = teamsRegion(answer.value.region ?? '');
      return region.ok
        ? region
        : err({ type: 'auth_failed', message: `The token helper gave the ${tier} token without a Teams region name, so no chat request was sent.`, code: tierCode(tier) });
    },
  };
};
