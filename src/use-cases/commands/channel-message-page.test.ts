import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import type { Command } from './command-types.ts';
import { CHANNEL_MESSAGES_TOP_CAP, CHANNEL_MESSAGES_TOP_OPTION, withChannelMessagesTopCap } from './channel-message-page.ts';

const graph = fakeGraphClient();

const spy = (): { calls: number; execute: Command['execute'] } => {
  const box = { calls: 0, execute: (async () => ok({})) as Command['execute'] };
  box.execute = async () => {
    box.calls += 1;
    return ok({ value: [] });
  };
  return box;
};

describe('the 50-message page cap on channel message reads', () => {
  it('lets a page of exactly the cap through to Graph', async () => {
    const inner = spy();
    const result = await withChannelMessagesTopCap(inner.execute)(graph, { top: String(CHANNEL_MESSAGES_TOP_CAP) });
    expect(result.ok).toBe(true);
    expect(inner.calls).toBe(1);
  });

  it('refuses a page one above the cap before calling Graph, naming the cap', async () => {
    const inner = spy();
    const result = await withChannelMessagesTopCap(inner.execute)(graph, { top: '51' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(result.error.message).toBe(
      '--top must be at most 50: Graph caps a page of channel messages at 50 and answers 400 above it. Continue through the `next:` footer with `next-page`.'
    );
    expect(inner.calls).toBe(0);
  });

  it('is silent when no --top is passed', async () => {
    const inner = spy();
    await withChannelMessagesTopCap(inner.execute)(graph, {});
    expect(inner.calls).toBe(1);
  });

  it('leaves a non-numeric --top to the schema validation behind it', async () => {
    const inner = spy();
    await withChannelMessagesTopCap(inner.execute)(graph, { top: 'many' });
    expect(inner.calls).toBe(1);
  });

  it('advertises the cap on the --top flag', () => {
    expect(CHANNEL_MESSAGES_TOP_CAP).toBe(50);
    expect(CHANNEL_MESSAGES_TOP_OPTION.name).toBe('top');
    expect(CHANNEL_MESSAGES_TOP_OPTION.key).toBe('top');
    expect(CHANNEL_MESSAGES_TOP_OPTION.required).toBe(false);
    expect(CHANNEL_MESSAGES_TOP_OPTION.description).toContain('50');
  });
});

describe('the helpers the markdown commands share', () => {
  it('turns a Graph nextLink into the relative path the client follows, and passes a missing link through', async () => {
    const { relativeGraphPath } = await import('./channel-message-page.ts');
    expect(relativeGraphPath('https://graph.microsoft.com/v1.0/teams/t/channels/c/messages?$skiptoken=x')).toBe('/teams/t/channels/c/messages?$skiptoken=x');
    expect(relativeGraphPath(undefined)).toBeUndefined();
  });

  it('builds the markdown envelope with a UTF-8 byte size and joins notes with a semicolon, omitting the field when there is none', async () => {
    const { markdownEnvelope } = await import('./channel-message-page.ts');
    expect(markdownEnvelope('héllo', [])).toEqual({ contentType: 'text/markdown', size: 6, text: 'héllo' });
    expect(markdownEnvelope('x', ['a', 'b'])).toEqual({ contentType: 'text/markdown', size: 1, text: 'x', note: 'a; b' });
  });
});

describe('what a Graph link must look like to be followed', () => {
  it('leaves a link that only contains the Graph origin somewhere inside it untouched', async () => {
    const { relativeGraphPath } = await import('./channel-message-page.ts');
    expect(relativeGraphPath('https://evil.example/https://graph.microsoft.com/v1.0/x')).toBe('https://evil.example/https://graph.microsoft.com/v1.0/x');
  });
});
