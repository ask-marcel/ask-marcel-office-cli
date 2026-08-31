/*
 * No-network guard for `bun test`.
 *
 * A unit test that reaches the real network is not a unit test: it passes or
 * fails on someone else's latency. Five in `auth.test.ts` POSTed a bogus
 * refresh token to `login.microsoftonline.com` and then passed or hit the 5s
 * limit depending on how fast AAD answered, so the suite's result moved between
 * 0, 3 and 5 failures on one unchanged tree (2026-08-31). They were fixed by
 * injection; this stops the next one being written at all.
 *
 * Wired in `bunfig.toml` under `[test] preload`, unlike
 * `scripts/coverage-preload.ts`, which is spawned only by
 * `scripts/check-coverage.ts` because it drags in heavy SDKs. This file imports
 * nothing and costs nothing, so every `bun test` run can afford it.
 *
 * There is deliberately NO escape hatch. Every IO boundary in `src/infra`
 * already takes an injectable `fetchFn` (`createGraphClient`,
 * `createAuthManagerFromApi`), so a test that wants a network answer has a
 * supported way to fake one. A test that genuinely must hit a live service
 * belongs in `scripts/qa-*`, which runs outside `bun test`.
 */

const realFetch = globalThis.fetch;

const blocked = (input: RequestInfo | URL): never => {
  const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
  throw new Error(
    `[no-network] a test tried to fetch ${url}. Tests must not touch the network. ` +
      `Inject a fake instead: \`createGraphClient(auth, fakeFetch)\`, or the trailing \`fetchFn\` argument of \`createAuthManagerFromApi\`. ` +
      `If neither fits, add an injection point — do not reach for the real service.`
  );
};

// Kept referenced so the real implementation is recoverable in a debugger and
// so this file never reads as if it simply deleted `fetch`.
Object.defineProperty(blocked, 'realFetch', { value: realFetch, enumerable: false });

globalThis.fetch = blocked as unknown as typeof fetch;
