# Security

Security guidance for Bun/TypeScript and Next.js code in the atelier style. Tuned for concrete, exploitable issues — not theoretical best-practice noise. Inspired by [anthropics/claude-code-security-review](https://github.com/anthropics/claude-code-security-review); the false-positive filter at the bottom is adapted with credit.

## Mental model | source to sink

Every security issue is a data flow problem:

1. **Source** | untrusted input. HTTP body / query / path, cookies, headers, user-uploaded files, third-party API responses, database rows whose writer is not trusted, deserialised JSON, browser-side `postMessage`, `window.location`, URL fragments.
2. **Sink** | an operation where untrusted data is dangerous. SQL query, shell command, filesystem path, HTTP request, HTML rendering, deserialisation, template compilation, `eval`, dynamic `import`, redirect destination, cookie value, log line (for secrets only).
3. **Checkpoint** | a branded type with a validating factory that sits between source and sink. Once data has type `Email`, `SafeUrl`, or `SanitizedHtml`, downstream code can trust it.

The rule is simple: **a source must cross a checkpoint before reaching a sink**. If you cannot point to the branded-type boundary, the flow is unsafe.

## Vulnerability categories to watch

Focus on HIGH and MEDIUM findings. Skip defence-in-depth hardening, theoretical races, and patterns where exploitation is not concretely possible.

### 1. Injection

- **SQL** | always use parameterised queries. Never string-interpolate untrusted values into SQL.
- **Command** | avoid `Bun.spawn` / `child_process` with untrusted arguments entirely. If unavoidable, use the array form, never a shell string.
- **Path traversal** | reject `..`, NUL bytes, and absolute paths in filenames from users. Resolve and verify with `path.resolve` + prefix check before opening the file.
- **SSRF** | outbound HTTP destinations chosen from user input need a host allowlist and a protocol allowlist. Only flag SSRF when the host or protocol is attacker-controlled; path-only SSRF is not a vulnerability.
- **Template / eval** | never pass untrusted strings to `eval`, `new Function`, `vm.runInNewContext`, or a template engine's raw compile API.

### 2. Cross-site scripting (XSS)

- **React / Next.js** | JSX auto-escapes. A React component is safe unless it uses `dangerouslySetInnerHTML`, renders `href={userInput}` with a `javascript:` URL, or passes untrusted HTML to a third-party renderer. Flag only those cases.
- **MDX** | content is compiled at build time. Untrusted MDX content (from users, not the repo) is dangerous and must not be rendered.
- **Response headers** | set `Content-Security-Policy` in `next.config.ts`. A strict CSP is the last line of defence.

### 3. Authentication & authorization

- **Always check on the server.** Client-side guards (disabled buttons, hidden routes) are UX, not security. The server must authenticate every request and authorise every resource access.
- **Session cookies** | `HttpOnly`, `Secure`, `SameSite=Lax` or stricter, set a sensible `Max-Age`.
- **JWT** | verify the signature with a fixed algorithm list (never accept `alg: none`). Verify `exp` and `iss`. Never trust the unverified payload.
- **Insecure Direct Object References** | always check that the caller owns / can access the resource they ask for, never assume the ID in the URL is theirs to use.

### 4. Crypto & secrets

- **Never roll your own crypto.** Use `crypto.subtle` (Web Crypto API, built into Bun) or a well-known library.
- **Random** | use `crypto.randomUUID()` or `crypto.getRandomValues()`. Never `Math.random()` for tokens, IDs, nonces.
- **Password hashing** | `argon2id` or `bcrypt` at a sensible cost factor. Never SHA-256 / MD5 / any fast hash.
- **TLS verification** | never disable certificate verification in production clients.
- **Secrets in code** | no API keys, tokens, passwords, or private keys in source. Ever. Load from `process.env` only.

### 5. Data exposure

- **Sensitive data in logs** | never log passwords, session tokens, refresh tokens, API keys, credit card numbers, or full PII records. Logging a URL or a user ID is fine; logging the `Authorization` header is not. Redact at the logger layer with a Winston format, not at every call site.
- **Error messages** | do not leak stack traces, internal paths, or SQL error messages to the client in production. Return a generic error and log the detail server-side.
- **Debug endpoints** | never expose `/debug`, `/__introspection`, or GraphQL introspection in production.

### 6. Deserialisation

- **JSON** | `JSON.parse` of untrusted input is safe for shape, but the result is still untrusted data — validate the shape with a branded-type factory or a schema (zod, valibot).
- **YAML** | `safe` mode only. Never `yaml.load` without a schema — YAML can instantiate arbitrary objects.
- **`node:vm` / `eval`** | never on untrusted input.

### 7. Supply chain

- `bun.lock` is committed. Review diffs on lockfile changes.
- Never install a package you have not heard of or cannot quickly verify on npm or its source repo.
- Pin versions in production images. Use `bun install --frozen-lockfile` in CI.

## Branded types for trust boundaries

The existing value-object pattern extends naturally to security. The factory is the checkpoint; the branded type is the proof that the check ran.

```ts
// URL validated for SSRF-safety (allowlisted hosts, only https)
const ALLOWED_HOSTS = new Set(['api.example.com', 'cdn.example.com']);

export type SafeUrl = string & { readonly __brand: 'SafeUrl' };

export const safeUrl = (value: string): SafeUrl => {
  const url = new URL(value);
  if (url.protocol !== 'https:') throw new Error('invalid SafeUrl.protocol');
  if (!ALLOWED_HOSTS.has(url.host)) throw new Error('invalid SafeUrl.host');
  return url.toString() as SafeUrl;
};

// only fetchers typed to accept SafeUrl can be called — grep catches every bypass
export const fetchJson = async (url: SafeUrl): Promise<unknown> => {
  const response = await fetch(url);
  return response.json();
};
```

```ts
// HTML that has been sanitised for rendering via dangerouslySetInnerHTML
import DOMPurify from 'isomorphic-dompurify';

export type SanitizedHtml = string & { readonly __brand: 'SanitizedHtml' };

export const sanitizedHtml = (raw: string): SanitizedHtml =>
  DOMPurify.sanitize(raw) as SanitizedHtml;
```

```ts
// Env var that must exist, must be non-empty, never leaks into logs
export type EnvVar = string & { readonly __brand: 'EnvVar' };

export const envVar = (name: string): EnvVar => {
  const value = process.env[name];
  if (value === undefined || value.length === 0) throw new Error(`Missing env var: ${name}`);
  return value as EnvVar;
};

// coerced + validated siblings — same gate at the env boundary, narrower output type
export const envNumber = (name: string): number => {
  const value = Number(envVar(name));
  if (!Number.isFinite(value)) throw new Error(`Env var ${name} is not a finite number`);
  return value;
};

export const envEnum = <T extends string>(name: string, allowed: readonly T[]): T => {
  const value = envVar(name);
  if (!allowed.includes(value as T)) throw new Error(`Env var ${name} must be one of: ${allowed.join(', ')}`);
  return value as T;
};

// one central config module; import from here, never read process.env directly elsewhere
export const config = {
  apiToken: envVar('API_TOKEN'),
  databaseUrl: envVar('DATABASE_URL'),
  logLevel: envEnum('LOG_LEVEL', ['error', 'warn', 'info', 'debug'] as const),
  port: envNumber('PORT'),
};
```

`envVar` brands the string because a raw `string` could bypass the non-empty check; the coerced siblings return their **natural** narrow types — `envNumber` a `number`, `envEnum` the literal union — which you wrap into a domain brand (`Port`, `TimeoutMs`) only where the value carries domain meaning (hard rule 12). A connection string consumed by a driver stays an `EnvVar`; a URL that will reach `fetch` is read with `safeUrl`, not `envVar`. When env outgrows a handful of vars or needs cross-field rules, parse it once at the composition root with a Zod schema instead — `const Env = z.object({ PORT: z.coerce.number().int().positive(), LOG_LEVEL: z.enum(['error', 'warn', 'info', 'debug']) }).parse(process.env)` — same principle either way: parse once, at the edge, fail loud at startup.

```ts
// file path that is guaranteed to live under a given root
import path from 'node:path';

export type SafePath = string & { readonly __brand: 'SafePath' };

export const safePath = (root: string, requested: string): SafePath => {
  if (requested.includes('\0')) throw new Error('invalid SafePath.nul');
  const resolved = path.resolve(root, requested);
  if (!resolved.startsWith(`${path.resolve(root)}${path.sep}`)) throw new Error('invalid SafePath.traversal');
  return resolved as SafePath;
};
```

The pattern is the same every time: one factory, one branded type, one truthful name. Callers cannot accidentally bypass the check because TypeScript rejects the raw `string`.

## Secrets and configuration

**Never commit:**
- `.env*` files (already gitignored).
- Service-account JSON (Firebase, GCP, AWS). Add `*-service-account*.json` to `.gitignore`.
- Private keys, certificates, SSH keys.
- API tokens, database passwords, OAuth client secrets, signing keys.
- Lockfile entries that contain credentials (some tools leak auth tokens into `.npmrc` or similar — audit before committing).

**Read only through a validated config module** (see `envVar` above). Never sprinkle `process.env.X` across the codebase.

**Logging discipline:**
```ts
// redact at the logger layer once; every call site benefits
// src/infra/logger.ts — adapter factory, wired once at composition (hard rule 4), never a module-level singleton
import { createLogger, format, transports } from 'winston';

const REDACTED_KEYS = new Set(['password', 'token', 'authorization', 'apiKey', 'secret']);

const redactFormat = format((info) => {
  for (const key of Object.keys(info)) {
    if (REDACTED_KEYS.has(key.toLowerCase())) info[key] = '[REDACTED]';
  }
  return info;
});

export const createWinstonLogger = (level: string): Logger =>
  createLogger({
    level,
    format: format.combine(redactFormat(), format.json()),
    transports: [new transports.Console()],
  });
```

Never interpolate secrets into log messages (`logger.info('token ' + token)`). Use structured logging and let the redactor do its job.

## Next.js specifics

- **`output: 'export'`** | static export means no runtime server code. There is no backend to defend in-package; every auth check must happen at your actual backend (Firebase, Cloud Run, etc.). Do not assume Next.js will filter anything at runtime.
- **CSP** | set `Content-Security-Policy` via `next.config.ts` `headers()`. Strict policy: no `'unsafe-inline'` for scripts, explicit allowlist for any CDN.
- **Images** | `images.unoptimized: true` (required for static export) means no server-side image sanitisation. Treat user-uploaded images as hostile; if you render them, do so only inside `<img>` with explicit dimensions, never via CSS `background-image` with user input.
- **MDX** | only render MDX sourced from your repository's `data/` folder. Never render user-submitted MDX.
- **Environment variables** | `NEXT_PUBLIC_*` variables ship to the browser. Never put secrets there.
- **Cookies** | set `HttpOnly` and `Secure` on any auth cookie. Use `SameSite=Lax` at minimum.

## Security review checklist (pre-merge)

Run through this before approving a change that touches trust boundaries:

1. **Sources identified** | for every new input, where does it come from? User, network, filesystem, environment, database?
2. **Sinks identified** | for every new output or operation, what consumes the data? SQL, shell, HTTP, filesystem, HTML, redirect, cookie, log?
3. **Checkpoints in place** | between each source and sink, is there a branded-type validating factory? If not, why not?
4. **authN/Z on the server** | if this is a backend change, is the check at the server boundary, not the client?
5. **No secrets leaked** | no new `console.log` of tokens; no new fields added to log objects without considering redaction; no secrets in source or in `NEXT_PUBLIC_*`.
6. **Dependencies reviewed** | did this PR change `bun.lock`? Are the new packages trustworthy?
7. **Error handling** | do errors avoid leaking stack traces or SQL messages to the client?

If the PR does not cross a trust boundary (pure refactor, domain logic, test changes), skip this checklist.

## Deep analysis before a release

For larger changes or a pre-release audit, install Anthropic's `/security-review` slash command from [anthropics/claude-code-security-review](https://github.com/anthropics/claude-code-security-review). Run it against the PR branch — it applies a stricter confidence threshold, a vetted false-positive filter, and a structured report format. It complements this reference rather than replacing it.

## False-positive filter (strict — adopt when reviewing)

When reviewing code for security issues, apply the same discipline the agent at `anthropics/claude-code-security-review` uses. Do NOT report:

1. Denial-of-Service or resource-exhaustion patterns. Only flag concrete, high-impact exploits.
2. A lack of hardening measures. Flag concrete vulnerabilities, not missing defence-in-depth.
3. Race conditions or timing attacks that are theoretical.
4. Outdated third-party libraries (that is a separate dependency-management concern).
5. Memory-safety issues in memory-safe languages.
6. Log spoofing, or logging of non-PII user input.
7. SSRF where only the path is attacker-controlled (host and protocol must be).
8. Including user-controlled content in AI prompts.
9. Regex injection or regex-DOS.
10. Findings in documentation (`.md` files, comments).
11. Lack of audit logs.
12. Open redirects, tab-nabbing, XS-Leaks, prototype pollution — unless extremely high confidence.
13. XSS in React or Angular components that do not use `dangerouslySetInnerHTML`, `bypassSecurityTrustHtml`, or similar explicit escape hatches.
14. GitHub Actions workflow issues unless a concrete untrusted-input path is demonstrated.
15. Client-side-only authN/Z gaps (client code is not a trust boundary).
16. Command injection in shell scripts not driven by untrusted input.
17. Logging URLs, UUIDs, or environment-variable values (treated as trusted).
18. Unit-test code.

**Confidence threshold** | only report a finding if you are at least 80% confident it is concretely exploitable. If you cannot describe the attack in one sentence, it is not a finding.

**Severity scale:**
- **HIGH** | directly exploitable, leads to RCE, data breach, auth bypass, privilege escalation.
- **MEDIUM** | exploitable under specific realistic conditions, significant impact.
- **LOW** | skip, unless the user has explicitly asked for a full defence-in-depth pass.

**Report format** (when producing a security review):

```markdown
# <category>: `<file>:<line>`

* Severity: High | Medium
* Description: <one sentence; what the vulnerability is>
* Exploit scenario: <concrete attack in 1-2 sentences>
* Recommendation: <specific fix; name the branded type or library>
```

Noise costs trust. Err on the side of not reporting.
