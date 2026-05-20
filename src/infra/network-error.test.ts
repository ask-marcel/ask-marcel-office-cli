import { describe, expect, it } from 'bun:test';
import {
  BINARY_TRANSFER_TIMEOUT_LABEL,
  BINARY_TRANSFER_TIMEOUT_MS,
  REQUEST_TIMEOUT_LABEL,
  REQUEST_TIMEOUT_MS,
  networkErrorMessage,
  timeoutLabelFor,
  timeoutMsFor,
} from './network-error.ts';

describe('networkErrorMessage — formats transport failures with call-site context', () => {
  it('reports a timeout against the JSON tier label when the underlying error is a TimeoutError', () => {
    const e = Object.assign(new Error('aborted'), { name: 'TimeoutError' });
    expect(networkErrorMessage(e, 'GET /me/messages', REQUEST_TIMEOUT_LABEL)).toBe(
      'request timed out after 60s (GET /me/messages) — transient; retry once before treating as permanent'
    );
  });

  it('reports a timeout against the binary tier label when the underlying error is a TimeoutError on a binary download', () => {
    const e = Object.assign(new Error('aborted'), { name: 'TimeoutError' });
    expect(networkErrorMessage(e, 'GET /drives/d1/items/i1/content', BINARY_TRANSFER_TIMEOUT_LABEL)).toBe(
      'request timed out after 5min (GET /drives/d1/items/i1/content) — transient; retry once before treating as permanent'
    );
  });

  it('surfaces the abort case distinctly from the timeout case so the LLM caller can tell user-cancel apart from deadline-hit', () => {
    const e = Object.assign(new Error('user cancel'), { name: 'AbortError' });
    expect(networkErrorMessage(e, 'GET /chats', REQUEST_TIMEOUT_LABEL)).toBe('request aborted (GET /chats) — transient; retry once before treating as permanent');
  });

  it('passes through a generic Error message when the failure is neither a timeout nor an abort', () => {
    const e = new Error('fetch failed');
    expect(networkErrorMessage(e, 'POST /search/query', REQUEST_TIMEOUT_LABEL)).toBe('fetch failed (POST /search/query) — transient; retry once before treating as permanent');
  });

  it('passes through a thrown plain string as the message body', () => {
    expect(networkErrorMessage('socket reset', 'GET /me', REQUEST_TIMEOUT_LABEL)).toBe('socket reset (GET /me) — transient; retry once before treating as permanent');
  });

  it('falls back to a stable string when the thrown value is neither an Error nor a string', () => {
    expect(networkErrorMessage({ weird: 'object' }, 'GET /me', REQUEST_TIMEOUT_LABEL)).toBe(
      'network request failed (GET /me) — transient; retry once before treating as permanent'
    );
  });
});

describe('timeoutLabelFor / timeoutMsFor — single source of truth for the two-tier timeout pairing', () => {
  it('returns the 60s pair for the JSON tier', () => {
    expect(timeoutLabelFor('json')).toBe(REQUEST_TIMEOUT_LABEL);
    expect(timeoutMsFor('json')).toBe(REQUEST_TIMEOUT_MS);
  });

  it('returns the 5min pair for the binary tier', () => {
    expect(timeoutLabelFor('binary')).toBe(BINARY_TRANSFER_TIMEOUT_LABEL);
    expect(timeoutMsFor('binary')).toBe(BINARY_TRANSFER_TIMEOUT_MS);
  });
});
