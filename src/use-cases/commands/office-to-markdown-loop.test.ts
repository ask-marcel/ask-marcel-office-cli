import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { officeToMarkdown } from './office-to-markdown.ts';

const htmlGraph = (html: string): ReturnType<typeof fakeGraphClient> => fakeGraphClient({ getBinary: async () => ok({ contentType: 'text/html', size: html.length, text: html }) });

describe('a Loop page whose HTML conversion is empty', () => {
  it('says so in a note instead of answering an empty body as if the page were empty', async () => {
    const result = await officeToMarkdown(htmlGraph(''), '/drives/d/items/i/content', 'notes.loop');
    if (!result.ok) throw new Error('expected ok');
    const envelope = result.value as { text: string; note?: string };
    expect(envelope.text).toBe('');
    expect(envelope.note).toContain('Loop converter lags the saves');
    expect(envelope.note).toContain('list-drive-item-versions');
  });

  it('adds nothing when the page renders', async () => {
    const result = await officeToMarkdown(htmlGraph('<h1>Agenda</h1><p>Budget</p>'), '/drives/d/items/i/content', 'notes.loop');
    if (!result.ok) throw new Error('expected ok');
    const envelope = result.value as { text: string; note?: string };
    expect(envelope.text).toContain('Agenda');
    expect(envelope.note).toBeUndefined();
  });
});
