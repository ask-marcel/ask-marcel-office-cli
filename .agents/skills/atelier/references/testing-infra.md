# Testing infra adapters

Infra adapters are the quarantine zone where thrown exceptions from third-party libraries become `Result<T, PortError>` values. Testing them is not testing the domain — it is testing the translation layer.

This reference is the companion to `references/testing.md`. The general "what is the unit, what is faked, no mocks" rules come from there; this file covers the patterns that apply specifically to infra-layer adapter tests.

There are three canonical patterns, one per dependency shape. Use the right pattern for the shape; do not mix.

- HTTP via `globalThis.fetch` → `installFetchMock` (per-test global swap with `afterEach().restore()`)
- External SDK → **dependency injection** with one of three sub-patterns: custom-fetch DI / two-constructor / sync-builder export
- Filesystem → real temp dirs, with `chmod` to trigger catch branches

Plus two cross-cutting bits at the end:

- The production-wiring smoke test that hits `createX(realDeps)` with placeholder credentials so the wiring line is covered without any network IO
- The fetch-mock ordering gotcha (silent failures) — most-specific-first, or `endsWith`

---

## 1. HTTP via `globalThis.fetch` → `installFetchMock`

For adapters that call `fetch` directly (Telegram, RSS fetcher, most HTTP-based adapters): swap `globalThis.fetch` via a helper that records every call and restores the real `fetch` in `afterEach`.

```ts
// src/test-helpers/fetch-mock.ts
export type FetchHandler = {
  readonly match: (url: string, init?: RequestInit) => boolean;
  readonly respond: (url: string, init?: RequestInit) => Response | Promise<Response>;
};

export type FetchMock = {
  readonly calls: ReadonlyArray<{ readonly url: string; readonly init?: RequestInit }>;
  readonly restore: () => void;
};

export const installFetchMock = (handlers: ReadonlyArray<FetchHandler>): FetchMock => {
  const calls: { url: string; init?: RequestInit }[] = [];
  const original = globalThis.fetch;
  globalThis.fetch = (async (input, init) => {
    const url = typeof input === 'string' ? input : input.url;
    calls.push({ url, init });
    const handler = handlers.find((h) => h.match(url, init));
    if (!handler) throw new Error(`fetch-mock: no handler for ${url}`);
    return handler.respond(url, init);
  }) as typeof fetch;
  return { calls, restore: () => { globalThis.fetch = original; } };
};
```

Used in a test:

```ts
describe('telegramHttp.send', () => {
  let mock: FetchMock;
  afterEach(() => mock?.restore());

  it('when the Telegram API returns 200, returns ok', async () => {
    mock = installFetchMock([
      {
        match: (url) => url.endsWith('/sendMessage'),
        respond: () => new Response(JSON.stringify({ ok: true, result: { message_id: 42 } })),
      },
    ]);
    const telegram = createTelegramHttp({ botToken: 'test' });

    const result = await telegram.send('chat-1', 'hello');

    expect(result.ok).toBe(true);
    expect(mock.calls).toHaveLength(1);
    expect(mock.calls[0].url).toContain('/sendMessage');
  });

  it('when the Telegram API returns 429, returns err rate-limited', async () => {
    mock = installFetchMock([
      {
        match: (url) => url.endsWith('/sendMessage'),
        respond: () => new Response('', { status: 429, headers: { 'retry-after': '5' } }),
      },
    ]);
    const telegram = createTelegramHttp({ botToken: 'test' });

    const result = await telegram.send('chat-1', 'hello');

    expect(result.ok).toBe(false);
    expect(!result.ok && result.error.kind).toBe('rate-limited');
  });
});
```

Assertions land on the returned `Result`, not on how `fetch` was called. The `calls` array is there for the occasional URL-shape check, not for interaction testing.

## 2. External SDK → dependency injection (three sub-patterns)

For adapters that import a third-party SDK, use **dependency injection**, never `mock.module`. `mock.module` is process-global: once set in one test file, the substitution leaks into every subsequent file the runner loads. `bun:test`'s `mock` namespace is banned outright (see `references/workflow.md`).

Three sub-patterns, in order of preference. Pick the first that applies.

### 2a. SDK accepts a custom `fetch` (preferred when available)

Many modern SDKs accept a custom `fetch` implementation as a constructor option — Vercel AI's `createGoogleGenerativeAI({ fetch })`, the OpenAI SDK's `new OpenAI({ fetch })`, the Anthropic SDK's `new Anthropic({ fetch })`. When integrating any new SDK, **check the docs for a custom-fetch option first**. If one exists, the whole adapter becomes end-to-end testable with zero SDK mocking.

Thread the fetch through your factory as an optional argument. Production omits it (the SDK falls back to `globalThis.fetch`); the test passes a fake fetch returning a canned response.

```ts
// src/infra/gemini-llm.ts
import { createGoogleGenerativeAI } from '@ai-sdk/google';
import { generateText, Output } from 'ai';

export const callGemini = (fetchImpl?: typeof globalThis.fetch): GenerateOutput => {
  const provider = createGoogleGenerativeAI(fetchImpl ? { fetch: fetchImpl } : {});
  return async ({ modelName, prompt, schema }) => {
    try {
      const result = await generateText({
        model: provider(modelName),
        prompt,
        output: Output.object({ schema }),
      });
      return ok(result.output);
    } catch (e) {
      return err({ kind: 'generate-failed', message: formatError(e) });
    }
  };
};

// Production wiring — no fetch passed; SDK uses globalThis.fetch.
export const createGeminiLlm = (): Llm => ({
  summarise: (text) => callGemini()({ modelName: 'gemini-2.5-flash', prompt: `...${text}`, schema }),
});
```

Test:

```ts
import { describe, expect, it } from 'bun:test';
import { callGemini } from './gemini-llm.ts';

describe('callGemini', () => {
  it('when the SDK returns the generated text, returns ok', async () => {
    const fakeFetch = (async () =>
      new Response(JSON.stringify({ candidates: [{ content: { parts: [{ text: '{"summary":"ok"}' }] } }] }))
    ) as unknown as typeof globalThis.fetch;

    const gen = callGemini(fakeFetch);
    const result = await gen({ modelName: 'gemini-1.5-flash', prompt: 'x', schema });

    expect(result.ok && result.value.summary).toBe('ok');
  });
});
```

**Cast note.** `typeof globalThis.fetch` includes `preconnect` (a Bun-specific extension on `fetch`, not part of the WHATWG Fetch standard). A bare arrow function cannot be assigned directly, and `as typeof globalThis.fetch` fails TypeScript's overlap heuristic. Use `as unknown as typeof globalThis.fetch` — the double-cast is load-bearing and `@typescript-eslint/no-unnecessary-type-assertion` will **not** flag it because the conversion truly spans an incompatible gap.

### 2b. No custom-fetch hook → two-constructor pattern

When the SDK has no fetch injection but its method surface is sliceable (`googleapis`, `firebase-admin`), split the SDK instantiation from the adapter logic. Every adapter exposes **two** exports from day one: `createX(realDeps)` (production wiring, one-liner) and `createXFromApi(api: XApi)` (the testable factory). `XApi` is a minimal type slice covering only the methods this adapter calls — never the full SDK type.

```ts
// src/infra/drive-google.ts

// 1) Minimal type slice — only the SDK surface this adapter actually calls.
export type DriveApi = {
  readonly files: {
    readonly copy: (params: { readonly fileId: string; readonly requestBody: { readonly name: string } }) =>
      Promise<{ readonly data: { readonly id: string } }>;
    readonly get: (params: { readonly fileId: string }) =>
      Promise<{ readonly data: { readonly name: string } }>;
  };
};

// 2) The testable factory — takes the API shape directly, returns the port.
export const createDriveFromApi = (api: DriveApi): Drive => ({
  copy: async (fileId, name) => {
    try {
      const res = await api.files.copy({ fileId, requestBody: { name } });
      return ok({ id: res.data.id });
    } catch (e) {
      return err({ kind: 'copy-failed', message: formatError(e) });
    }
  },
  getName: async (fileId) => {
    try {
      const res = await api.files.get({ fileId });
      return ok(res.data.name);
    } catch (e) {
      return err({ kind: 'not-found', message: formatError(e) });
    }
  },
});

// 3) The production wiring — one-liner that instantiates the real SDK.
//    `as unknown as DriveApi` is permitted HERE ONLY, because real SDK types
//    are overloaded and rarely structurally match a hand-written slice.
export const createGoogleDrive = (auth: GoogleAuth): Drive =>
  createDriveFromApi(google.drive({ version: 'v3', auth: auth.client }) as unknown as DriveApi);
```

The test imports `createDriveFromApi` and passes an in-memory object that satisfies the slice.

```ts
import { describe, expect, it } from 'bun:test';
import { createDriveFromApi, type DriveApi } from './drive-google.ts';

describe('driveGoogle.copy', () => {
  it('when the SDK returns an id, returns ok with that id', async () => {
    const api: DriveApi = {
      files: {
        copy: async () => ({ data: { id: 'copied-123' } }),
        get: async () => ({ data: { name: 'ignored' } }),
      },
    };
    const drive = createDriveFromApi(api);

    const result = await drive.copy('src-1', 'My Copy');

    expect(result.ok && result.value).toEqual({ id: 'copied-123' });
  });
});
```

If strict lint flags the `as unknown as XApi` cast as unnecessary, the SDK's real type structurally matched the slice — drop the cast. Keep it only when the compiler genuinely needs it.

### Anti-pattern: `XApi` shaped like the port

The two-constructor pattern only buys testability when `XApi` slices the **SDK's** surface. If `XApi` is shaped like the **port**, the seam is in the wrong place and `createX` stays untestable.

**Bad — `BrowserAuthApi` is a clone of the port:**

```ts
// src/infra/browser-auth.ts
export type BrowserAuthApi = {
  readonly acquireToken: (scopes: ReadonlyArray<string>) => Promise<string>;
  readonly close: () => Promise<void>;
};

export const createBrowserAuthFromApi = (api: BrowserAuthApi): BrowserAuth => ({
  acquire: async (scopes) => {
    try { return ok({ token: await api.acquireToken(scopes) }); }
    catch (e) { return err({ kind: 'acquire-failed', message: formatError(e) }); }
  },
  close: api.close,
});

// What does createBrowserAuth(realDeps) actually look like?
// It has to call Playwright — but BrowserAuthApi doesn't say HOW.
// All the real logic (launchPersistentContext, polling for the auth callback,
// extracting the cookie) lives inside createBrowserAuth and isn't reachable
// from createBrowserAuthFromApi at all.
```

`createBrowserAuthFromApi` is a tautological pass-through. The test you can write against it proves nothing, and the production wiring stays a black box.

**Good — `PlaywrightApi` slices the SDK's real surface:**

```ts
// src/infra/browser-auth.ts
export type PlaywrightApi = {
  readonly launchPersistentContext: (
    userDataDir: string,
    options: { readonly headless: boolean }
  ) => Promise<{
    readonly newPage: () => Promise<{
      readonly goto: (url: string) => Promise<void>;
      readonly waitForURL: (pattern: RegExp, options: { readonly timeout: number }) => Promise<void>;
      readonly url: () => string;
    }>;
    readonly cookies: () => Promise<ReadonlyArray<{ readonly name: string; readonly value: string }>>;
    readonly close: () => Promise<void>;
  }>;
};

export const createBrowserAuthFromApi = (api: PlaywrightApi, config: BrowserAuthConfig): BrowserAuth => ({
  acquire: async (scopes) => {
    try {
      const ctx = await api.launchPersistentContext(config.userDataDir, { headless: false });
      const page = await ctx.newPage();
      await page.goto(buildAuthUrl(scopes));
      await page.waitForURL(/code=/, { timeout: config.navigationTimeoutMs });
      const cookies = await ctx.cookies();
      return ok({ token: extractTokenFromCookies(cookies) });
    } catch (e) {
      return err({ kind: 'acquire-failed', message: formatError(e) });
    }
  },
  // ...
});

export const createBrowserAuth = (config: BrowserAuthConfig): BrowserAuth =>
  createBrowserAuthFromApi(playwright.chromium as unknown as PlaywrightApi, config);
```

Now the test passes a fake `PlaywrightApi` that returns a stub `ctx` with stub pages and stub cookies — and the **real** orchestration logic inside `createBrowserAuthFromApi` (build URL, wait for redirect, extract token) is exercised. `createBrowserAuth` becomes the genuine one-line wiring it should be, covered by the production-wiring smoke test below.

**The diagnostic.** Look at your `XApi` and your port side-by-side. If the method names are nearly identical and the parameter shapes match 1:1, you've made a port clone. The right slice usually has SDK-flavoured names (`launchPersistentContext`, `files.copy`, `chat.completions.create`) and SDK-flavoured option bags — because that's what the production code actually calls.

### 2c. Sync constructor → export the private builder

Some SDKs are sync constructors that return an object whose methods you call later (`new TwitterApi(creds).v1.tweet(...)`, `new MongoClient(url).db(...)`). Slicing is awkward because the real object shape is used throughout the adapter. Instead, export the otherwise-private builder factory so the test can invoke it directly with placeholder credentials. The constructor itself is offline; no network IO happens until a method is called.

```ts
// src/infra/twitter-api.ts
export const buildRealClient = (creds: TwitterCreds): TwitterApi =>
  new TwitterApi({ appKey: creds.appKey, appSecret: creds.appSecret, accessToken: creds.token, accessSecret: creds.secret });

export const createTwitterApi = (creds: TwitterCreds): TwitterPort => {
  const client = buildRealClient(creds);
  return {
    postTweet: async (text) => {
      try {
        const res = await client.v2.tweet(text);
        return ok({ id: res.data.id });
      } catch (e) {
        return err({ kind: 'post-failed', message: formatError(e) });
      }
    },
  };
};
```

Test:

```ts
const placeholderCreds = { appKey: 'x', appSecret: 'x', token: 'x', secret: 'x' };
const client = buildRealClient(placeholderCreds);
expect(client.v1).toBeDefined();
expect(client.v2).toBeDefined();
```

The constructor runs — the wiring line executes and covers — but no network method is called.

### Production-wiring smoke test (2b and 2c)

`createX(realDeps)` is the production wiring line. Without a test that calls it, the coverage tool reports 0% on that line and the 80% gate for `src/infra/**` will fail. Every adapter that uses pattern 2b or 2c gets a one-describe-block smoke test that calls the production factory with placeholder auth / credentials and asserts the returned object has the port's method shape.

```ts
describe('createGoogleDrive (production wiring smoke)', () => {
  it('returns a Drive port with the expected method shape', () => {
    const drive = createGoogleDrive({ client: {} as never });
    expect(typeof drive.copy).toBe('function');
    expect(typeof drive.getName).toBe('function');
  });
});
```

The smoke test does **two** jobs at once:

1. It exercises the wiring line (the `createXFromApi(realSdk(...))` call) so the per-tier coverage gate passes without launching the real SDK.
2. It pins the module as reachable, satisfying the `coverage-preload.ts` invariant — every infra file must be importable from the preload chain, and every infra file must have a test that touches it. The smoke test is the cheapest way to do both.

Methods are asserted as `typeof === 'function'`, not invoked. Invoking would require either a real Playwright browser, a real Google Drive client, or a fake — which would defeat the point. The whole purpose is "prove the wiring compiles and produces the right shape, without doing anything else."

Pattern 2a (custom-fetch DI) does not need a separate smoke test because the production wiring is itself exercised end-to-end by passing `fakeFetch`.

### Configurable durations for IO loops with deadlines

Adapters that retry, poll, or wait for external state always need a deadline. In production, the deadline matches user expectations (a 5-minute auth callback window, a 30-second LLM timeout, a 2-second between-poll delay). In tests, the same deadlines would crawl the suite to a halt — and `setTimeout` global swaps only help when the duration is genuinely "fire immediately".

The pattern: every duration the adapter cares about is a field on a `Config` record passed into the factory. Production wiring fills it with real values; tests pass tiny values.

```ts
// src/infra/browser-auth.ts
export type BrowserAuthConfig = {
  readonly userDataDir: string;
  readonly initialSettleMs: number;       // wait after page load before reading state
  readonly pollIntervalMs: number;        // between checks for the redirect
  readonly pollDeadlineMs: number;        // total time before giving up
  readonly navigationTimeoutMs: number;   // single page navigation
};

// Production defaults. `Partial<>` so the caller only overrides what differs.
export const productionBrowserAuthConfig = (overrides: Partial<BrowserAuthConfig> = {}): BrowserAuthConfig => ({
  userDataDir: '.auth-cache',
  initialSettleMs: 1_000,
  pollIntervalMs: 500,
  pollDeadlineMs: 5 * 60_000,
  navigationTimeoutMs: 30_000,
  ...overrides,
});

export const createBrowserAuth = (config: BrowserAuthConfig): BrowserAuth => /* ... */;
```

In production:

```ts
// src/composition/build-deps.ts
const auth = createBrowserAuth(productionBrowserAuthConfig());
```

In tests:

```ts
const fastConfig: BrowserAuthConfig = {
  userDataDir: tmp,
  initialSettleMs: 1,
  pollIntervalMs: 1,
  pollDeadlineMs: 50,
  navigationTimeoutMs: 50,
};
const auth = createBrowserAuthFromApi(fakePlaywright, fastConfig);
```

**Rule.** Any IO loop with a deadline (retry, poll, fetch-with-timeout, queue-drain, debounce) names every duration in a `Config` record and accepts it through the factory. The defaults factory (`productionXConfig`) holds the real values; tests construct `fastConfig` literals. This pulls the test runtime down by orders of magnitude without touching `setTimeout` at all.

**Rule**: every new SDK adapter picks one of 2a, 2b, or 2c, and the pattern you chose must be visible from the exports. For 2b: both `createX` and `createXFromApi` exported. For 2c: `buildRealClient` also exported. For 2a: the `fetchImpl?` parameter visible in the signature. If none of these is true, the adapter has no test seam and the next change will reach for `mock.module`.

### Sub-pattern: swapping a global (e.g. `setTimeout`) per-test

When an adapter calls a global like `setTimeout` — for real retry delays, say — tests will crawl unless the global is swapped. **Do not** reach for `mock.module`. Swap the global in `beforeAll` / `afterAll`; the scope is the test file, not the whole process.

```ts
import { afterAll, beforeAll, describe, expect, it } from 'bun:test';

describe('instagramGraph.pollUntilReady', () => {
  let originalSetTimeout: typeof globalThis.setTimeout;
  beforeAll(() => {
    originalSetTimeout = globalThis.setTimeout;
    globalThis.setTimeout = ((fn: () => void) => {
      fn();
      return 0;
    }) as typeof globalThis.setTimeout;
  });
  afterAll(() => { globalThis.setTimeout = originalSetTimeout; });

  it('polls and returns ok once ready', async () => {
    // ... test runs fast; setTimeout is a no-op
  });
});
```

Same principle as `installFetchMock`: the swap is bounded by a lifecycle hook that always restores the original. No process-global surprise.

## 3. Filesystem → real temp dirs

For adapters that read or write the filesystem (`token-store-fs`, `prompt-loader-fs`, `Bun.file` wrappers): use real temp directories. Mocking `Bun.file` or `node:fs` is fragile and misses real edge cases (permission errors, malformed files, disappearing symlinks). A temp directory is fast (microseconds) and exercises the real semantics.

```ts
import { afterEach, beforeEach, describe, expect, it } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createPromptLoaderFs } from './prompt-loader-fs.ts';

describe('promptLoaderFs', () => {
  let tmp: string;
  beforeEach(() => { tmp = mkdtempSync(join(tmpdir(), 'prompt-loader-')); });
  afterEach(() => rmSync(tmp, { recursive: true, force: true }));

  it('when a prompt file exists, returns ok with its content', async () => {
    writeFileSync(join(tmp, 'summary.md'), 'hello world');
    const loader = createPromptLoaderFs({ root: tmp });

    const result = await loader.load('summary');

    expect(result.ok && result.value).toBe('hello world');
  });

  it('when a prompt file is missing, returns err not-found', async () => {
    const loader = createPromptLoaderFs({ root: tmp });

    const result = await loader.load('absent');

    expect(!result.ok && result.error.kind).toBe('not-found');
  });
});
```

### Triggering read/write catch blocks (chmod, not a directory)

To hit the `catch` branch that guards a `Bun.file(path).text()` or `Bun.write(path, ...)` call, the instinct is to pass a directory path — but `Bun.file(dir).exists()` returns **`false`** for directories, so the read routes to the `not-found` branch instead of throwing. The real way to force a thrown exception is `chmod` on a real file (or directory) and restore in `finally`.

```ts
import { chmodSync } from 'node:fs';

it('when the file is unreadable (chmod 0000), returns err read-failed', async () => {
  const path = join(tmp, 'locked.md');
  await Bun.write(path, 'content');
  chmodSync(path, 0o000);
  try {
    const result = await createPromptLoaderFs({ root: tmp }).load('locked');

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.kind).toBe('read-failed');
  } finally {
    chmodSync(path, 0o600); // so afterEach's rmSync can clean up
  }
});

it('when the directory is read-only (chmod 0500), Bun.write returns err write-failed', async () => {
  chmodSync(tmp, 0o500);
  try {
    const store = createTokenStoreFs({ dir: tmp });
    const result = await store.save('token-1', 'secret');

    expect(!result.ok && result.error.kind).toBe('write-failed');
  } finally {
    chmodSync(tmp, 0o700);
  }
});
```

The restore in `finally` is mandatory — without it, `afterEach`'s `rmSync` cannot remove the locked file and the next test starts dirty.

**Platform note.** `chmod` is a Unix primitive; on Windows it is a silent no-op. These tests exercise the catch branch on Linux and macOS CI runners only. If the project ever runs Windows CI, skip these with `if (process.platform !== 'win32') { ... }`.

## Fetch-mock handler ordering (silent gotcha)

Fetch-mock handlers are checked in array order, first match wins. A broad match like `url.includes('/IG123/media')` will also match `/IG123/media_publish` — Instagram tests failed with `"network-failed"` instead of `"publish-failed"` because of this. Fix: use `url.endsWith('/IG123/media')` for exact suffix matching, or put the more specific handler (`/media_publish`) **first** in the handlers array. Silent failures are the worst kind; prefer `endsWith` by default.

