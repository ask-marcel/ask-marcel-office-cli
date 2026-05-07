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

describe('presenter output', () => {
  it('renders a JSON line to stdout and logs an info event', async () => {
    const logger = createLoggerFake();
    const out = await captureStream('stdout', () => render({ status: 'authenticated' }, logger));
    expect(out.trim()).toBe(JSON.stringify({ status: 'authenticated' }));
    expect(logger.calls.some((c) => c.event === 'output_rendered')).toBe(true);
  });

  it('renders an error envelope to stdout (not stderr) as exactly one line so jq can parse it from a single stream', async () => {
    const out = await captureStream('stdout', () => renderError('Authentication cancelled'));
    expect(out.trim()).toBe(JSON.stringify({ error: 'Authentication cancelled' }));
    expect(out.split('\n').filter((line) => line.length > 0)).toHaveLength(1);
  });

  it('writes nothing to stderr when an error is rendered (errors live on stdout in v1)', async () => {
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
    const reparsed = JSON.parse(trimmed) as { value: ReadonlyArray<{ summary: string; nested: { description: string } }> };
    expect(reparsed.value[0]?.summary).toBe(payload);
    expect(reparsed.value[0]?.nested.description).toBe(payload);
  });

  it('escapes U+2028 and U+2029 line/paragraph separators so the output round-trips through JSON.parse', async () => {
    const logger = createLoggerFake();
    const data = { line: 'a b', paragraph: 'c d' };
    const out = await captureStream('stdout', () => render(data, logger));
    const trimmed = out.replace(/\n$/, '');
    const reparsed = JSON.parse(trimmed) as { line: string; paragraph: string };
    expect(reparsed.line).toBe('a b');
    expect(reparsed.paragraph).toBe('c d');
  });
});
