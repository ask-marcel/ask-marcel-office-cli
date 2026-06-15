/*
 * Fetch-mock test helper for atelier infra adapter tests.
 *
 * Swaps globalThis.fetch with a handler-driven stub that records every call
 * and restores the real fetch in afterEach. Used by adapters that call
 * globalThis.fetch directly (Telegram, RSS fetcher, HTTP-based adapters).
 *
 * Usage:
 *
 *   import { afterEach } from 'bun:test';
 *   import { installFetchMock } from '../test-helpers/fetch-mock.ts';
 *
 *   let mock: ReturnType<typeof installFetchMock> | undefined;
 *   afterEach(() => mock?.restore());
 *
 *   mock = installFetchMock([
 *     { match: (url) => url.endsWith('/sendMessage'),
 *       respond: () => new Response(JSON.stringify({ ok: true })) },
 *   ]);
 *
 * IMPORTANT: handlers are checked in order, first match wins. Put the
 * more-specific matcher first (e.g. /api/foo_publish before /api/foo),
 * or use url.endsWith(...) for exact path-suffix matching. A broad
 * url.includes(...) will match more URLs than you expect.
 *
 * See skills/atelier/references/testing.md (Testing infra adapters, pattern 1).
 */

// FetchInput / FetchInit are derived from the global `fetch` signature so the
// helper compiles without requiring the `DOM` lib in tsconfig.
type FetchInput = Parameters<typeof fetch>[0];
type FetchInit = Parameters<typeof fetch>[1];

export type FetchHandler = {
  readonly match: (url: string, init: FetchInit) => boolean;
  readonly respond: (url: string, init: FetchInit) => Response | Promise<Response>;
};

export type FetchMockCall = {
  readonly url: string;
  readonly init: FetchInit;
};

export type FetchMock = {
  readonly calls: ReadonlyArray<FetchMockCall>;
  readonly restore: () => void;
};

const urlOf = (input: FetchInput): string => {
  if (typeof input === 'string') return input;
  if (input instanceof URL) return input.toString();
  return input.url;
};

export const installFetchMock = (handlers: ReadonlyArray<FetchHandler>): FetchMock => {
  const calls: FetchMockCall[] = [];
  const original = globalThis.fetch;
  globalThis.fetch = (async (input: FetchInput, init: FetchInit): Promise<Response> => {
    const url = urlOf(input);
    calls.push({ url, init });
    const handler = handlers.find((h) => h.match(url, init));
    if (!handler) throw new Error(`fetch-mock: no handler matched ${url}`);
    return handler.respond(url, init);
  }) as typeof fetch;
  return {
    calls,
    restore: (): void => {
      globalThis.fetch = original;
    },
  };
};
