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

describe('presenter output (v1 envelope)', () => {
  it('wraps a successful render in { ok: true, data } and logs an info event', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ status: 'authenticated' }, logger));
    expect(JSON.parse(out.trim())).toEqual({ ok: true, data: { status: 'authenticated' } });
    expect(logger.calls.some((c) => c.event === 'output_rendered')).toBe(true);
  });

  it('lifts @odata.nextLink to the top level and removes it from data', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'm1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' };
    const out = await captureStream('stdout', () => render(data, logger));
    const parsed = JSON.parse(out.trim()) as { ok: boolean; data: Record<string, unknown>; nextLink?: string };
    expect(parsed.ok).toBe(true);
    expect(parsed.nextLink).toBe('https://graph.microsoft.com/v1.0/me/messages?$skip=10');
    expect(parsed.data).toEqual({ value: [{ id: 'm1' }] });
  });

  it('lifts @odata.count to the top level and removes it from data', async () => {
    const logger = createLoggerFake();
    const data = { value: [{ id: 'm1' }, { id: 'm2' }], '@odata.count': 42 };
    const out = await captureStream('stdout', () => render(data, logger));
    const parsed = JSON.parse(out.trim()) as { ok: boolean; data: Record<string, unknown>; count?: number };
    expect(parsed.ok).toBe(true);
    expect(parsed.count).toBe(42);
    expect(parsed.data).toEqual({ value: [{ id: 'm1' }, { id: 'm2' }] });
  });

  it('omits nextLink and count when neither @odata field is present', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ id: 'me' }, logger));
    const parsed = JSON.parse(out.trim()) as Record<string, unknown>;
    expect(Object.keys(parsed).toSorted()).toEqual(['data', 'ok']);
  });

  it('wraps a non-object data value (string / number / null) in { ok: true, data } unchanged', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render('plain string', logger));
    expect(JSON.parse(out.trim())).toEqual({ ok: true, data: 'plain string' });
  });

  it('wraps an array data value in { ok: true, data } without lifting @odata.* keys (arrays cannot host them)', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render([1, 2, 3], logger));
    expect(JSON.parse(out.trim())).toEqual({ ok: true, data: [1, 2, 3] });
  });

  it('wraps an error message in { ok: false, error } and writes to stdout (not stderr)', async () => {
    const out = await captureStream('stdout', () => renderError('Authentication cancelled'));
    expect(JSON.parse(out.trim())).toEqual({ ok: false, error: 'Authentication cancelled' });
  });

  it('writes nothing to stderr when an error is rendered', async () => {
    const out = await captureStream('stderr', () => renderError('Boom'));
    expect(out).toBe('');
  });

  it('escapes every U+0000..U+001F control character in string leaves so the output round-trips through JSON.parse', async () => {
    const logger = createLoggerFake();
    let payload = '';
    for (let cp = 0; cp <= 0x1f; cp += 1) payload += String.fromCharCode(cp);
    const data = { value: [{ summary: payload, nested: { description: payload } }] };
    const out = await captureStream('stdout', () => render(data, logger));
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
    const out = await captureStream('stdout', () => render(data, logger));
    const trimmed = out.replace(/\n$/, '');
    const parsed = JSON.parse(trimmed) as { data: { line: string; paragraph: string } };
    expect(parsed.data.line).toBe('a b');
    expect(parsed.data.paragraph).toBe('c d');
  });
});
