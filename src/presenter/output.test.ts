import { describe, expect, it } from 'bun:test';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { render, renderError } from './output.ts';

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

  it('lifts @odata.deltaLink to the top level alongside nextLink so resumption tokens sit at the envelope level (audit v1.0.0 §4)', async () => {
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
    const user = { id: '0c1d', displayName: 'Vincent Delacourt', mail: 'vincent@example.com' };
    const out = await captureStream('stdout', () => render(user, logger, 'text'));
    expect(out).toBe('id: 0c1d\ndisplayName: Vincent Delacourt\nmail: vincent@example.com\n');
    expect(logger.calls.some((c) => c.event === 'output_rendered')).toBe(true);
  });

  it('renders a nested object by indenting the sub-keys two spaces under their parent', async () => {
    const logger = createLoggerFake();
    const data = { user: { displayName: 'Vincent', mail: 'vincent@example.com' }, primaryDriveId: 'b!abc' };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('user:\n  displayName: Vincent\n  mail: vincent@example.com\nprimaryDriveId: b!abc\n');
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

  it('appends a footer line carrying the next-page cursor when a listing has @odata.nextLink', async () => {
    const logger = createLoggerFake();
    const data = {
      value: [{ id: 'm1', subject: 'hi' }],
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=10',
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('id: m1\nsubject: hi\n\n--- next: https://graph.microsoft.com/v1.0/me/messages?$skip=10\n');
  });

  it('packs nextLink, deltaLink, and count side-by-side into a single footer separated by middle dots', async () => {
    const logger = createLoggerFake();
    const data = {
      value: [{ id: 'e1' }],
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/events?$skip=10',
      '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta?$dt=X',
      '@odata.count': 47,
    };
    const out = await captureStream('stdout', () => render(data, logger, 'text'));
    expect(out).toBe('id: e1\n\n--- next: https://graph.microsoft.com/v1.0/me/events?$skip=10 · delta: https://graph.microsoft.com/v1.0/me/events/delta?$dt=X · count: 47\n');
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
    expect(out).toBe('(no items)\n\n--- next: https://graph.microsoft.com/v1.0/me/messages?$skip=10\n');
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

  it('renders an error as a single "error: <message>" line without the JSON envelope so an LLM can read the failure directly', async () => {
    const out = await captureStream('stdout', () => renderError('missing scope: Calendars.Read', 'text'));
    expect(out).toBe('error: missing scope: Calendars.Read\n');
  });

  it('writes nothing to stderr when a text-mode error is rendered (single-stream contract)', async () => {
    const out = await captureStream('stderr', () => renderError('Boom', 'text'));
    expect(out).toBe('');
  });
});
