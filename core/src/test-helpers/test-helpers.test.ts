import { afterEach, describe, expect, it } from 'bun:test';
import { installFetchMock } from './fetch-mock.ts';
import { createLoggerFake } from './logger-fake.ts';

describe('logger-fake', () => {
  it('records warn calls', () => {
    const logger = createLoggerFake();
    logger.warn('test_warning', { key: 'value' });
    expect(logger.calls).toHaveLength(1);
    expect(logger.calls[0]).toEqual({ level: 'warn', event: 'test_warning', meta: { key: 'value' } });
  });

  it('records info and error calls', () => {
    const logger = createLoggerFake();
    logger.info('test_info');
    logger.error('test_error', { detail: 'oops' });
    expect(logger.calls).toHaveLength(2);
    expect(logger.calls[0].level).toBe('info');
    expect(logger.calls[1].level).toBe('error');
  });
});

describe('fetch-mock', () => {
  afterEach(() => {
    if (globalThis.fetch !== originalFetch) globalThis.fetch = originalFetch;
  });

  const originalFetch = globalThis.fetch;

  it('records calls with string URL', () => {
    const mock = installFetchMock([{ match: () => true, respond: () => new Response('ok') }]);
    void fetch('https://example.com/api');
    expect(mock.calls).toHaveLength(1);
    expect(mock.calls[0].url).toBe('https://example.com/api');
    mock.restore();
  });

  it('records calls with URL object', () => {
    const mock = installFetchMock([{ match: () => true, respond: () => new Response('ok') }]);
    void fetch(new URL('https://example.com/api'));
    expect(mock.calls).toHaveLength(1);
    expect(mock.calls[0].url).toBe('https://example.com/api');
    mock.restore();
  });
});
