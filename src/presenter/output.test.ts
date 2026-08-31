import { describe, expect, it } from 'bun:test';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { renderTextOutput } from './output-text.ts';
import { render, renderError } from './output.ts';
import type { RenderContext } from './render-to-string.ts';
import { renderToString } from './render-to-string.ts';

const captureStream = async (stream: 'stdout' | 'stderr', run: () => void | Promise<void>): Promise<string> => {
  const target = process[stream];
  const original = target.write.bind(target);
  let captured = '';
  const swap = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  target.write = swap;
  try {
    await run();
  } finally {
    target.write = original;
  }
  return captured;
};

describe('presenter output — JSON envelope (opt-in via --output json)', () => {
  it('wraps a successful render in { ok: true, data } and logs an info event', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ status: 'authenticated' }, logger, 'json'));
    expect(JSON.parse(out.trim())).toEqual({ ok: true, data: { status: 'authenticated' } });
    expect(logger.calls.some((c) => c.event === 'output_rendered')).toBe(true);
  });

  it('lifts @odata.nextLink to the top level and removes it from data', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'm1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' };
    const out = await captureStream('stdout', () => render(data, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: boolean; data: Record<string, unknown>; nextLink?: string };
    expect(parsed.ok).toBe(true);
    expect(parsed.nextLink).toBe('https://graph.microsoft.com/v1.0/me/messages?$skip=10');
    expect(parsed.data).toEqual({ value: [{ id: 'm1' }] });
  });

  it('lifts @odata.count to the top level and removes it from data', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'm1' }, { id: 'm2' }], '@odata.count': 42 };
    const out = await captureStream('stdout', () => render(data, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: boolean; data: Record<string, unknown>; count?: number };
    expect(parsed.ok).toBe(true);
    expect(parsed.count).toBe(42);
    expect(parsed.data).toEqual({ value: [{ id: 'm1' }, { id: 'm2' }] });
  });

  it('lifts @odata.deltaLink to the top level alongside nextLink so resumption tokens sit at the envelope level', async () => {
    const logger = createLoggerFake();
    const data = {
      value: [{ id: 'e1', subject: 'standup' }],
      '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta()?$deltatoken=ABC',
    };
    const out = await captureStream('stdout', () => render(data, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: boolean; data: Record<string, unknown>; deltaLink?: string };
    expect(parsed.ok).toBe(true);
    expect(parsed.deltaLink).toBe('https://graph.microsoft.com/v1.0/me/events/delta()?$deltatoken=ABC');
    expect(parsed.data).toEqual({ value: [{ id: 'e1', subject: 'standup' }] });
  });

  it('canonicalises a %24-encoded $ in the hoisted nextLink so `/me/people?%24top=2&%24skip=2` surfaces as `$top`/`$skip` (Graph encodes the operator; re-encoding it would break next-page)', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'p1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/people?%24top=2&%24skip=2' };
    const out = await captureStream('stdout', () => render(data, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { nextLink?: string };
    expect(parsed.nextLink).toBe('https://graph.microsoft.com/v1.0/me/people?$top=2&$skip=2');
  });

  it('canonicalises a %24-encoded $ in the hoisted deltaLink the same way as nextLink', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'e1' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta()?%24deltatoken=ABC' };
    const out = await captureStream('stdout', () => render(data, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { deltaLink?: string };
    expect(parsed.deltaLink).toBe('https://graph.microsoft.com/v1.0/me/events/delta()?$deltatoken=ABC');
  });

  it('leaves non-$ percent-escapes inside an opaque skiptoken VALUE untouched (only %24 is decoded — %2B/%2F/%3D encode literal token bytes that must survive a round-trip)', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'm1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?%24skiptoken=aB%2Bc%2Fd%3D' };
    const out = await captureStream('stdout', () => render(data, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { nextLink?: string };
    expect(parsed.nextLink).toBe('https://graph.microsoft.com/v1.0/me/messages?$skiptoken=aB%2Bc%2Fd%3D');
  });

  it('omits nextLink and count when neither @odata field is present', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ id: 'me' }, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as Record<string, unknown>;
    expect(Object.keys(parsed).toSorted((a, b) => a.localeCompare(b))).toEqual(['data', 'ok']);
  });

  it('wraps a non-object data value (string / number / null) in { ok: true, data } unchanged', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render('plain string', logger, 'json'));
    expect(JSON.parse(out.trim())).toEqual({ ok: true, data: 'plain string' });
  });

  it('wraps an array data value in { ok: true, data } without lifting @odata.* keys (arrays cannot host them)', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render([1, 2, 3], logger, 'json'));
    expect(JSON.parse(out.trim())).toEqual({ ok: true, data: [1, 2, 3] });
  });

  it('wraps an error message in { ok: false, error } and writes to stdout (not stderr)', async () => {
    const out = await captureStream('stdout', () => renderError('Authentication cancelled', 'json'));
    expect(JSON.parse(out.trim())).toEqual({ ok: false, error: 'Authentication cancelled' });
  });

  // — structured-error path. The hint table maps the
  // high-frequency Graph errors (InvalidIdMalformed, MissingScope, …) to a
  // one-line remedy + a source classifier ('graph' | 'cli' | 'validation').
  // The envelope is additive — old consumers keying on `error: string` still
  // work; new consumers branch on `hint` / `source`.
  it('attaches { hint, source } to the JSON envelope when the errorCode matches a known Graph error (e.g. ErrorInvalidIdMalformed)', async () => {
    const out = await captureStream('stdout', () => renderError('ErrorInvalidIdMalformed: Id is malformed.', 'json', 'ErrorInvalidIdMalformed'));
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string; errorCode: string; hint?: string; source?: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.errorCode).toBe('ErrorInvalidIdMalformed');
    expect(parsed.hint).toContain('Source IDs from a sibling');
    expect(parsed.source).toBe('graph');
  });

  it('omits hint/source from the JSON envelope when nothing in the hint table matches AND no explicit source was supplied (preserves the historical 3-arg back-compat shape)', async () => {
    const out = await captureStream('stdout', () => renderError('Some weird new failure mode', 'json', 'WeirdNewCode'));
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string; errorCode: string; hint?: string; source?: string };
    expect(parsed.hint).toBeUndefined();
    expect(parsed.source).toBeUndefined();
  });

  // — envelope-shape stability fix. The user
  // reported 5 bare-Graph error codes that lacked `source` even though every
  // CLI-side validation error carries it. The fix: callers that know the
  // failure category (cli.ts maps the GraphError discriminated-union type to
  // an `ErrorSource`) pass it explicitly; the presenter stamps it even when
  // no hint rule matched. End-result envelope is `{ok, error, errorCode?,
  // hint?, source?}` where only `hint` is conditional.
  it('emits `source` from the explicit 4th-arg even when no hint rule matched — closes the bare-Graph-error envelope asymmetry', async () => {
    const out = await captureStream('stdout', () => renderError('Some weird new Graph failure', 'json', 'WeirdNewCode', 'graph'));
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string; errorCode: string; hint?: string; source?: string };
    expect(parsed.hint).toBeUndefined();
    expect(parsed.source).toBe('graph');
  });

  it('surfaces the Retry-After interval in the JSON envelope so a throttled CLI consumer can honor Graph backoff', async () => {
    const out = await captureStream('stdout', () => renderError('TooManyRequests: Too many requests', 'json', 'TooManyRequests', 'graph', 120));
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string; errorCode: string; retryAfterSeconds?: number };
    expect(parsed.ok).toBe(false);
    expect(parsed.retryAfterSeconds).toBe(120);
  });

  it('surfaces Retry-After as a `retryAfter: Ns` line in text output', async () => {
    const out = await captureStream('stdout', () => renderError('TooManyRequests: slow down', 'text', 'TooManyRequests', 'graph', 30));
    expect(out).toContain('retryAfter: 30s');
  });

  it('omits retryAfterSeconds from the envelope when no interval was supplied (most errors carry none)', async () => {
    const out = await captureStream('stdout', () => renderError('itemNotFound: not found', 'json', 'itemNotFound', 'graph'));
    const parsed = JSON.parse(out.trim()) as { ok: false; retryAfterSeconds?: number };
    expect(parsed.retryAfterSeconds).toBeUndefined();
  });

  it('keeps a Retry-After of 0 in the envelope (0 is an immediate-retry hint, not absence)', async () => {
    const out = await captureStream('stdout', () => renderError('TooManyRequests: slow down', 'json', 'TooManyRequests', 'graph', 0));
    const parsed = JSON.parse(out.trim()) as { retryAfterSeconds?: number };
    expect(parsed.retryAfterSeconds).toBe(0);
  });

  it('hint-table source wins over the explicit 4th-arg when both are available (curated rule beats discriminator-derived fallback)', async () => {
    // Caller asserts `source: 'graph'` but the rule table classifies
    // `cli_reject_search_with_filter` as `cli` — the hint's source wins.
    const out = await captureStream('stdout', () =>
      renderError(
        '--filter is incompatible with $search on /me/messages — Graph rejects the combination with `SearchWithFilter`.',
        'json',
        'cli_reject_search_with_filter',
        'graph'
      )
    );
    const parsed = JSON.parse(out.trim()) as { ok: false; source?: string };
    expect(parsed.source).toBe('cli');
  });

  it('text-mode also stamps the explicit `source:` line when no hint matched (envelope symmetry across both formats)', async () => {
    const out = await captureStream('stdout', () => renderError('Some weird new Graph failure', 'text', 'WeirdNewCode', 'graph'));
    const lines = out.trim().split('\n');
    expect(lines[0]).toBe('error: Some weird new Graph failure');
    expect(lines[1]).toBe('source: graph');
    expect(lines.length).toBe(2);
  });

  it('appends `hint:` and `source:` lines to text-mode errors so an LLM matching on `error:` still works but ALSO gets the remedy', async () => {
    const out = await captureStream('stdout', () => renderError('ErrorInvalidIdMalformed: Id is malformed.', 'text', 'ErrorInvalidIdMalformed'));
    const lines = out.trim().split('\n');
    expect(lines[0]).toBe('error: ErrorInvalidIdMalformed: Id is malformed.');
    expect(lines[1]).toMatch(/^hint: The ID you passed isn't valid for this endpoint/);
    expect(lines[2]).toBe('source: graph');
  });

  it('text-mode errors keep the single-line shape when nothing matches the hint table AND no explicit source was supplied (back-compat for the 3-arg form)', async () => {
    const out = await captureStream('stdout', () => renderError('Some weird new failure mode', 'text', 'WeirdNewCode'));
    expect(out).toBe('error: Some weird new failure mode\n');
  });

  // — sizeHint when the rendered envelope crosses
  // 50 KB. The remedy is command-aware (2026-07-23): a render with no command
  // behind it — the manifest, a login summary, a logout status — claims no
  // flag-level remedy at all, because it cannot know one applies.
  it('adds a `sizeHint` field to the JSON envelope when the rendered envelope exceeds 50 KB, and a render with no command behind it recommends no flag it cannot vouch for', async () => {
    const logger = createLoggerFake();
    // 60 KB of dummy data — well over the 50 KB threshold and the hint
    // itself can't push the borderline case over (additive guard in
    // renderJson measures the initial envelope size, not the post-hint one).
    const big = { value: Array.from({ length: 600 }, (_, i) => ({ id: `item-${i}`, payload: 'x'.repeat(100) })) };
    const out = await captureStream('stdout', () => render(big, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: true; sizeHint?: string };
    expect(parsed.sizeHint).toBeDefined();
    expect(parsed.sizeHint).toContain('> 50 KB threshold');
    // The banner used to promise `--output-path` "works on every command" and
    // to name endpoints that ignore --select/--top. Both were dead ends on the
    // commands that print it most (search-all-files, microsoft-search-query
    // advertise only --query), so a context-free render now promises neither.
    expect(parsed.sizeHint).not.toContain('--output-path');
    expect(parsed.sizeHint).not.toContain('> out.json');
  });

  it('omits `sizeHint` when the rendered envelope fits inside 50 KB (no warning churn on small responses)', async () => {
    const logger = createLoggerFake();
    const small = { value: [{ id: '1', subject: 'hi' }] };
    const out = await captureStream('stdout', () => render(small, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: true; sizeHint?: string };
    expect(parsed.sizeHint).toBeUndefined();
  });

  it('text-mode renders prepend a `sizeHint:` line above the body when the body exceeds 50 KB so the LLM sees the warning before scrolling', async () => {
    const logger = createLoggerFake();
    const big = { value: Array.from({ length: 600 }, (_, i) => ({ id: `item-${i}`, payload: 'x'.repeat(100) })) };
    const out = await captureStream('stdout', () => render(big, logger, 'text'));
    expect(out.startsWith('sizeHint: Response is ')).toBe(true);
    expect(out).toContain('> 50 KB threshold');
  });
});

// 2026-07-23 bug report: the banner promised `--output-path` as a "universal
// remedy (works on every command)", but the commands that trip it most often
// are plain-JSON ones that REFUSE the flag — so the caller burned a call to
// learn the banner was wrong. The remedy is now derived from the command that
// produced the payload and the surface the caller is on.
describe('presenter output — the oversized-response banner only names remedies the caller can actually use', () => {
  const oversized = { value: Array.from({ length: 600 }, (_, i) => ({ id: `item-${i}`, payload: 'x'.repeat(100) })) };
  const hintOf = (context: RenderContext): string => {
    const parsed = JSON.parse(renderToString(oversized, 'json', context).trim()) as { sizeHint?: string };
    return parsed.sizeHint ?? '';
  };

  it('an oversized download from the terminal is told to write the body to a file with --output-path', () => {
    const hint = hintOf({ commandName: 'download-drive-item-content', producesBytes: true, supportsSelect: false, supportsTop: false, surface: 'cli' });
    expect(hint).toContain('--output-path <file>');
    expect(hint).not.toContain('> out.json');
  });

  it('the same download over MCP is told to set the outputPath param, since an MCP client cannot pass a --flag', () => {
    const hint = hintOf({ commandName: 'download-drive-item-content', producesBytes: true, supportsSelect: false, supportsTop: false, surface: 'mcp' });
    expect(hint).toContain('`outputPath`');
    expect(hint).not.toContain('--output-path');
  });

  it('an oversized search from the terminal is told to redirect to a file, with the command named in the example', () => {
    const hint = hintOf({ commandName: 'search-all-files', producesBytes: false, supportsSelect: false, supportsTop: false, surface: 'cli' });
    expect(hint).toContain('ask-marcel-office search-all-files ... > out.json');
    expect(hint).toContain('refused');
  });

  it('the same search over MCP is never told to redirect, because an MCP client has no shell', () => {
    const hint = hintOf({ commandName: 'search-all-files', producesBytes: false, supportsSelect: false, supportsTop: false, surface: 'mcp' });
    expect(hint).not.toContain('> out.json');
    expect(hint).toContain('refused');
  });

  it('a command that advertises neither --select nor --top is told to narrow the query text instead of naming flags it lacks', () => {
    const hint = hintOf({ commandName: 'microsoft-search-query', producesBytes: false, supportsSelect: false, supportsTop: false, surface: 'cli' });
    expect(hint).toContain('narrow the query text');
    expect(hint).not.toContain('--select');
    expect(hint).not.toContain('--top');
  });

  it('a command that advertises both slimming flags has both named', () => {
    const hint = hintOf({ commandName: 'list-mail-folder-messages', producesBytes: false, supportsSelect: true, supportsTop: true, surface: 'cli' });
    expect(hint).toContain('--select');
    expect(hint).toContain('--top');
  });

  it('a command that advertises only --select is not told to pass --top', () => {
    const hint = hintOf({ commandName: 'list-planner-plans', producesBytes: false, supportsSelect: true, supportsTop: false, surface: 'cli' });
    expect(hint).toContain('--select');
    expect(hint).not.toContain('--top');
  });
});

describe('presenter output — JSON envelope (opt-in via --output json)', () => {
  // v1.4.0 audit #4: when `--select` is given all unknown field names,
  // Graph silently drops every field and returns `value: [{@odata.etag},
  // {@odata.etag}, ...]` — N entries that look "empty" once the etag is
  // stripped. The presenter surfaces a `selectHint` to flag the likely
  // typo. Distinct from `sizeHint` (which fires on byte count).
  it('adds a `selectHint` to the JSON envelope when `value[]` has entries but each entry is empty after stripping @odata.etag (likely bogus `--select`)', async () => {
    const logger = createLoggerFake();
    const bogusSelect = {
      value: [{ '@odata.etag': 'W/"abc1"' }, { '@odata.etag': 'W/"abc2"' }, { '@odata.etag': 'W/"abc3"' }],
    };
    const out = await captureStream('stdout', () => render(bogusSelect, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: true; selectHint?: string };
    expect(parsed.selectHint).toBeDefined();
    expect(parsed.selectHint).toContain('--select');
    expect(parsed.selectHint).toContain('responseShape');
  });

  it('omits `selectHint` when `value[]` is legitimately empty (zero matches — distinguishable from bogus-select because there are no entries to be empty)', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ value: [] }, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: true; selectHint?: string };
    expect(parsed.selectHint).toBeUndefined();
  });

  it('omits `selectHint` when entries carry any non-etag field (real data)', async () => {
    const logger = createLoggerFake();
    const realData = {
      value: [
        { id: '1', subject: 'a' },
        { id: '2', subject: 'b' },
      ],
    };
    const out = await captureStream('stdout', () => render(realData, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: true; selectHint?: string };
    expect(parsed.selectHint).toBeUndefined();
  });

  it('omits `selectHint` when the response has no `value[]` and carries real fields (single-resource GET — the most common shape, `get-current-user` etc.)', async () => {
    const logger = createLoggerFake();
    const single = { id: 'u1', displayName: 'Alice' };
    const out = await captureStream('stdout', () => render(single, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: true; selectHint?: string };
    expect(parsed.selectHint).toBeUndefined();
  });

  // The audit's specific case: `get-current-user --select aaaaaaaa,bbbbbbbb`
  // returns `{ "@odata.context": "..." }` — a single-resource GET with no
  // non-metadata keys. Same trap as the collection-shape variant, different
  // wire shape.
  it('also fires `selectHint` for the single-resource-GET bogus-select shape — top-level object with only @odata.* keys (e.g. `get-current-user --select bogus` returns just `{@odata.context}`)', async () => {
    const logger = createLoggerFake();
    const bogusSingle = { '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#users(aaaaaaaa,bbbbbbbb)/$entity' };
    const out = await captureStream('stdout', () => render(bogusSingle, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: true; selectHint?: string };
    expect(parsed.selectHint).toBeDefined();
    expect(parsed.selectHint).toContain('--select');
  });

  it('omits `selectHint` for the response with @odata.context AND real fields (the normal happy-path single-resource shape — context is always returned by Graph)', async () => {
    const logger = createLoggerFake();
    const normalSingle = { '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#users/$entity', id: 'u1', displayName: 'Alice' };
    const out = await captureStream('stdout', () => render(normalSingle, logger, 'json'));
    const parsed = JSON.parse(out.trim()) as { ok: true; selectHint?: string };
    expect(parsed.selectHint).toBeUndefined();
  });

  it('text-mode also surfaces the bogus-select warning as a `selectHint:` prelude line', async () => {
    const logger = createLoggerFake();
    const bogusSelect = { value: [{ '@odata.etag': 'W/"x"' }, { '@odata.etag': 'W/"y"' }] };
    const out = await captureStream('stdout', () => render(bogusSelect, logger, 'text'));
    expect(out.startsWith('selectHint: ')).toBe(true);
    expect(out).toContain('--select');
  });

  it('writes nothing to stderr when an error is rendered', async () => {
    const out = await captureStream('stderr', () => renderError('Boom', 'json'));
    expect(out).toBe('');
  });

  it('escapes every U+0000..U+001F control character in string leaves so the output round-trips through JSON.parse', async () => {
    const logger = createLoggerFake();
    let payload = '';
    for (let cp = 0; cp <= 0x1f; cp += 1) payload += String.fromCharCode(cp);
    const data = { value: [{ summary: payload, nested: { description: payload } }] };
    const out = await captureStream('stdout', () => render(data, logger, 'json'));
    const trimmed = out.replace(/\n$/, '');
    expect(trimmed.includes('\n')).toBe(false);
    expect(trimmed.includes('\t')).toBe(false);
    expect(trimmed.includes('\r')).toBe(false);
    const parsed = JSON.parse(trimmed) as { data: { value: ReadonlyArray<{ summary: string; nested: { description: string } }> } };
    expect(parsed.data.value[0]?.summary).toBe(payload);
    expect(parsed.data.value[0]?.nested.description).toBe(payload);
  });

  it('escapes U+2028 and U+2029 line/paragraph separators so the output round-trips through JSON.parse', async () => {
    const logger = createLoggerFake();
    const data = { line: 'a b', paragraph: 'c d' };
    const out = await captureStream('stdout', () => render(data, logger, 'json'));
    const trimmed = out.replace(/\n$/, '');
    const parsed = JSON.parse(trimmed) as { data: { line: string; paragraph: string } };
    expect(parsed.data.line).toBe('a b');
    expect(parsed.data.paragraph).toBe('c d');
  });
});

describe('presenter output — text format (default for LLM consumers)', () => {
  it('renders a single user profile as YAML-ish key:value lines an LLM can scan without parsing', async () => {
    const logger = createLoggerFake();
    const user = { id: '0c1d', displayName: 'Jordan Avery', mail: 'jordan.avery@example.com' };
    const out = await captureStream('stdout', () => render(user, logger, 'text'));
    expect(out).toBe('id: 0c1d\ndisplayName: Jordan Avery\nmail: jordan.avery@example.com\n');
    expect(logger.calls.some((c) => c.event === 'output_rendered')).toBe(true);
  });

  it('renders a nested object by indenting the sub-keys two spaces under their parent', async () => {
    const logger = createLoggerFake();
    const data = { user: { displayName: 'Jordan', mail: 'jordan.avery@example.com' }, primaryDriveId: 'b!abc' };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('user:\n  displayName: Jordan\n  mail: jordan.avery@example.com\nprimaryDriveId: b!abc\n');
  });

  it('renders a Graph collection { value: [...] } as one YAML-ish item block per record separated by blank lines', async () => {
    const logger = createLoggerFake();
    const data = {
      value: [
        { id: 'm1', subject: 'Re: Q2 planning', from: 'alice@example.com' },
        { id: 'm2', subject: 'Lunch?', from: 'bob@example.com' },
      ],
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('id: m1\nsubject: Re: Q2 planning\nfrom: alice@example.com\n\nid: m2\nsubject: Lunch?\nfrom: bob@example.com\n');
  });

  it('renders the next-page cursor as a whole ready-to-run command with a SINGLE-quoted URL (so $ / & survive a shell paste), not a bare URL', async () => {
    const logger = createLoggerFake();
    const data = {
      value: [{ id: 'm1', subject: 'hi' }],
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=10',
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe("id: m1\nsubject: hi\n\n--- next: ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/me/messages?$skip=10'\n");
  });

  it('packs the nextLink + deltaLink run-commands and the raw count side-by-side into a single footer separated by middle dots', async () => {
    const logger = createLoggerFake();
    const data = {
      value: [{ id: 'e1' }],
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/events?$skip=10',
      '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta?$dt=X',
      '@odata.count': 47,
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe(
      "id: e1\n\n--- next: ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/me/events?$skip=10' · delta: ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/me/events/delta?$dt=X' · count: 47\n"
    );
  });

  it('canonicalises a %24-encoded $ in the text footer cursor so the copy-pasteable `next:` value matches the `next-page` $-form', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'p1', displayName: 'Robin Chen' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/people?%24top=2&%24skip=2' };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe("id: p1\ndisplayName: Robin Chen\n\n--- next: ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/me/people?$top=2&$skip=2'\n");
  });

  it('emits no footer line when a listing carries no pagination cursors and no count', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'm1', subject: 'only one' }] };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('id: m1\nsubject: only one\n');
  });

  it('renders an empty Graph collection as a "(no items)" line so the LLM does not misread silence as a crash', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ value: [] }, logger, 'text'));
    expect(out).toBe('(no items)\n');
  });

  it('keeps the cursor footer even when the empty listing carries a nextLink (next page might have items)', async () => {
    const logger = createLoggerFake();
    const data = { value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe("(no items)\n\n--- next: ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/me/messages?$skip=10'\n");
  });

  // A partner-tenant listing signs page 1 with a guest token, and the cursor
  // Graph hands back carries no tenant. The footer's whole promise is that the
  // line can be copied verbatim, so it has to carry the flag that re-signs the
  // next page — otherwise the copy lands on `401 invalidAudienceUri`.
  it('appends --tenant-id to the next: command when the originating call carried one, so the copied line resolves in the partner tenant', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'f1', name: 'Report.docx' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/children?$skip=10' };
    const context: RenderContext = {
      commandName: 'list-folder-files',
      producesBytes: false,
      supportsSelect: false,
      supportsTop: false,
      surface: 'cli',
      tenantId: '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04',
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text', context));
    expect(out).toBe(
      "id: f1\nname: Report.docx\n\n--- next: ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/children?$skip=10' --tenant-id 6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04\n"
    );
  });

  it('appends --tenant-id to the delta: command too, since a delta cursor is re-signed through the same next-page path', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'f1' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/drives/d1/root/delta?$dt=X' };
    const context: RenderContext = {
      commandName: 'list-folder-files',
      producesBytes: false,
      supportsSelect: false,
      supportsTop: false,
      surface: 'cli',
      tenantId: '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04',
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text', context));
    expect(out).toBe(
      "id: f1\n\n--- delta: ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/drives/d1/root/delta?$dt=X' --tenant-id 6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04\n"
    );
  });

  it('leaves the footer untouched when the call carried no tenant, which is every home-tenant listing', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'm1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' };
    const context: RenderContext = { commandName: 'list-mail-messages', producesBytes: false, supportsSelect: true, supportsTop: true, surface: 'cli' };
    const out = await captureStream('stdout', () => render(data, logger, 'text', context));
    expect(out).toBe("id: m1\n\n--- next: ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/me/messages?$skip=10'\n");
    expect(out).not.toContain('--tenant-id');
  });

  // MCP renders the same text footer, and an MCP client cannot run a shell
  // line — but it reads the flag names off it to build its own params, so the
  // tenant has to be visible there too.
  it('carries the tenant on the MCP surface as well, where the footer is what the agent reads to build its next call', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'f1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/children?$skip=10' };
    const context: RenderContext = {
      commandName: 'list-folder-files',
      producesBytes: false,
      supportsSelect: false,
      supportsTop: false,
      surface: 'mcp',
      tenantId: '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04',
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text', context));
    expect(out).toContain('--tenant-id 6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04');
  });

  it('renders a 20,000-item aggregated collection with a plain count sibling (search-all-files shape) as flat blocks with the >50KB sizeHint prepended', async () => {
    const logger = createLoggerFake();
    const items = Array.from({ length: 20_000 }, (_, i) => ({ id: `item-${i}`, name: `File ${i}.docx`, webUrl: `https://contoso.sharepoint.com/f/${i}` }));
    const data = { value: items, count: 20_000 };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    // The sizeHint was previously unreachable in text mode for payloads this
    // large: the renderer crashed before the size check ever ran.
    expect(out.startsWith('sizeHint: ')).toBe(true);
    expect(out).toContain('\nid: item-0\nname: File 0.docx\n');
    expect(out.endsWith('id: item-19999\nname: File 19999.docx\nwebUrl: https://contoso.sharepoint.com/f/19999\n\n--- count: 20000\n')).toBe(true);
  });

  it('renders a 600,000-record array under a record key without a JavaScriptCore spread-argument stack overflow (search-all-files-scale payloads)', () => {
    // Sized above the measured JSC failure point for `arr.push(...big)`
    // (~400K-1M spread arguments depending on stack depth, Bun 1.3.x): the
    // old spread-push rendering threw `Maximum call stack size exceeded`
    // here. renderTextOutput is used directly so the 7 MB body skips the
    // captureStream indirection.
    const hits = Array.from({ length: 600_000 }, (_, i) => ({ id: `x${i}` }));
    const out = renderTextOutput({ hits });
    // 1 `hits:` header + 600,000 item lines + 599,999 blank separators + trailing newline
    expect(out.split('\n').length).toBe(1_200_001);
    expect(out.startsWith('hits:\n  id: x0\n')).toBe(true);
  });

  it('hoists a scalar sibling into the footer next to a hoisted @odata.count cursor (both land in the same `---` line)', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'm1' }], '@odata.count': 47, truncated: true };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('id: m1\n\n--- count: 47 · truncated: true\n');
  });

  it('renders an empty collection that carries scalar siblings as "(no items)" plus the hoisted footer', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ value: [], count: 0 }, logger, 'text'));
    expect(out).toBe('(no items)\n\n--- count: 0\n');
  });

  it('hoists plain scalar siblings of a value[] collection (count, truncated) into the footer instead of nesting the items under a `value:` tree', async () => {
    const logger = createLoggerFake();
    const data = {
      value: [
        { id: 'f1', name: 'a.docx' },
        { id: 'f2', name: 'b.pptx' },
      ],
      count: 2,
      truncated: true,
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('id: f1\nname: a.docx\n\nid: f2\nname: b.pptx\n\n--- count: 2 · truncated: true\n');
  });

  it('keeps the record-tree rendering when a value[] envelope carries a non-scalar sibling (nothing to hoist safely)', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'f1' }], stats: { pages: 3 } };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('value:\n  id: f1\nstats:\n  pages: 3\n');
  });

  it('prints the markdown body raw with no envelope when the command returns a text/markdown payload (convert-mail-to-markdown family)', async () => {
    const logger = createLoggerFake();
    const data = { contentType: 'text/markdown', size: 12, text: '# Hello\n\nbody.' };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('# Hello\n\nbody.\n');
  });

  it('also prints text/plain payloads raw (e.g. download-drive-item-as-markdown returning text/plain)', async () => {
    const logger = createLoggerFake();
    const data = { contentType: 'text/plain', size: 5, text: 'hello' };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('hello\n');
  });

  it('replaces inline base64 with a "use --output-path" hint so the LLM does not pull a multi-MB blob through stdout', async () => {
    const logger = createLoggerFake();
    const data = { contentType: 'application/pdf', size: 12345, base64: 'JVBERi0…' };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('binary: application/pdf, 12345 bytes — use --output-path to save\n');
  });

  it('renders a savedTo envelope as ordinary key:value lines after --output-path has consumed the inline bytes', async () => {
    const logger = createLoggerFake();
    const data = { contentType: 'application/pdf', size: 12345, savedTo: '/work/test/may-deck.pdf' };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('contentType: application/pdf\nsize: 12345\nsavedTo: /work/test/may-deck.pdf\n');
  });

  it('replaces an unsaved media array with a one-line "use --output-dir" hint in text output (base64 stays out of stdout)', async () => {
    const logger = createLoggerFake();
    const data = {
      count: 2,
      media: [
        { path: 'ppt/media/image1.png', contentType: 'image/png', sizeBytes: 2048, base64: 'iVBOR…' },
        { path: 'word/media/photo.jpeg', contentType: 'image/jpeg', sizeBytes: 4096, base64: '/9j/4…' },
      ],
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('2 image(s), 6 KB total — use --output-dir <dir> to extract them to disk (base64 omitted from text output; add --output json to inline it)\n');
  });

  it('renders a saved media array (savedTo, no base64) as ordinary key:value lines', async () => {
    const logger = createLoggerFake();
    const data = { count: 1, media: [{ path: 'ppt/media/image1.png', contentType: 'image/png', sizeBytes: 2048, savedTo: '/work/imgs/image1.png' }] };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toContain('savedTo: /work/imgs/image1.png');
    expect(out).not.toContain('use --output-dir');
  });

  it('inlines a flat array of primitive scope strings on the same line so a 30-item array stays one line tall', async () => {
    const logger = createLoggerFake();
    const data = { audience: 'graph', scopes: ['Mail.Read', 'Calendars.Read', 'Files.Read'] };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('audience: graph\nscopes: [Mail.Read, Calendars.Read, Files.Read]\n');
  });

  it('expands an array-of-records under a parent key into one item block per record', async () => {
    const logger = createLoggerFake();
    const data = {
      todoLists: [
        { id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' },
        { id: 'l2', displayName: 'Shopping' },
      ],
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('todoLists:\n  id: l1\n  displayName: Tasks\n  wellknownListName: defaultList\n\n  id: l2\n  displayName: Shopping\n');
  });

  it('renders a top-level string primitive as the string followed by a newline', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render('plain string', logger, 'text'));
    expect(out).toBe('plain string\n');
  });

  it('renders a top-level array of primitives as one value per line', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render([1, 2, 3], logger, 'text'));
    expect(out).toBe('1\n2\n3\n');
  });

  it('renders a single status record (login/logout/update success) as one key:value line', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ status: 'authenticated' }, logger, 'text'));
    expect(out).toBe('status: authenticated\n');
  });

  it('renders an unrecognised error as a single "error: <message>" line (no hint/source appended when the hint table does not match)', async () => {
    const out = await captureStream('stdout', () => renderError('some unmapped failure with no recognisable pattern', 'text'));
    expect(out).toBe('error: some unmapped failure with no recognisable pattern\n');
  });

  it('writes nothing to stderr when a text-mode error is rendered (single-stream contract)', async () => {
    const out = await captureStream('stderr', () => renderError('Boom', 'text'));
    expect(out).toBe('');
  });
});
