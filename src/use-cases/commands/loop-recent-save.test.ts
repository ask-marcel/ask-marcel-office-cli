import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const minutesAgo = (minutes: number): string => new Date(Date.now() - minutes * 60_000).toISOString();

const render = async (name: string, html: string, lastModifiedDateTime?: string): Promise<{ text: string; note?: string }> => {
  const command = commands['download-drive-item-as-markdown'];
  if (!command) throw new Error('download-drive-item-as-markdown is not registered');
  const graph = fakeGraphClient({
    get: async () => ok({ name, ...(lastModifiedDateTime === undefined ? {} : { lastModifiedDateTime }) }),
    getBinary: async () => ok({ contentType: 'text/html', size: html.length, text: html }),
  });
  const result = await command.execute(graph, { driveId: 'd1', itemId: 'i1' });
  if (!result.ok) throw new Error(result.error.message);
  return result.value as { text: string; note?: string };
};

const RECENT =
  "This page was saved in the last 30 minutes, and Graph's Loop render can trail the latest saves: if it looks stale, run this again in a while; `list-drive-item-versions` shows when it was saved.";

describe('a Loop page saved moments ago', () => {
  it('says the render may trail the latest saves', async () => {
    const envelope = await render('meeting-notes.loop', '<h1>Agenda</h1>', minutesAgo(5));
    expect(envelope.text).toContain('Agenda');
    expect(envelope.note).toBe(RECENT);
  });

  it('adds nothing for a page saved an hour ago, a page with no save time, or a recent file that is not a Loop page', async () => {
    expect((await render('meeting-notes.loop', '<h1>Agenda</h1>', minutesAgo(60))).note).toBeUndefined();
    expect((await render('meeting-notes.loop', '<h1>Agenda</h1>')).note).toBeUndefined();
    expect((await render('page.html', '<h1>Agenda</h1>', minutesAgo(5))).note).toBeUndefined();
  });

  it('keeps the empty-render note when the body is empty', async () => {
    expect((await render('meeting-notes.loop', '', minutesAgo(5))).note).toContain('Graph returned no HTML for this page');
  });
});

describe('a Loop page Graph fails to convert', () => {
  it('passes the failure on unchanged', async () => {
    const command = commands['download-drive-item-as-markdown'];
    if (!command) throw new Error('download-drive-item-as-markdown is not registered');
    const failure = { type: 'api_error' as const, status: 406, message: 'Sandbox_InputFormatNotSupported: not supported' };
    const graph = fakeGraphClient({
      get: async () => ok({ name: 'notes.loop', lastModifiedDateTime: minutesAgo(5) }),
      getBinary: async () => ({ ok: false as const, error: failure }),
    });
    const result = await command.execute(graph, { driveId: 'd1', itemId: 'i1' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('Sandbox_InputFormatNotSupported');
  });
});
