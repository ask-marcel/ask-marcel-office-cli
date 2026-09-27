import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { commands } from './index.ts';

type Answer = { readonly contentType: string; readonly size: number; readonly text: string; readonly note?: string };

const convert = async (html: string, params: Record<string, string> = {}): Promise<Answer> => {
  const command = commands['convert-local-file-to-markdown'];
  if (command?.executeLocal === undefined) throw new Error('convert-local-file-to-markdown is not a local command');
  const fs = createFileSystemFake();
  fs.seed('/work/page.html', html);
  const result = await command.executeLocal(fs, { path: '/work/page.html', ...params });
  if (!result.ok) throw new Error(result.error.message);
  return result.value as Answer;
};

// A "Save as Web Page" export: one wide table, about 1.2 MB of markup.
const bigExport = (): string => {
  const rows = Array.from({ length: 15000 }, (_unused, i) => `<tr><td>row ${i}</td><td>North region</td><td>${i * 3}</td><td>on track</td></tr>`).join('');
  return `<html><head><title>Export</title></head><body><table><tr><th>Line</th><th>Region</th><th>Units</th><th>Status</th></tr>${rows}</table></body></html>`;
};

const PNG = `data:image/png;base64,${'iVBORw0KGgo'.repeat(2000)}`;

describe('an HTML page converted to markdown', () => {
  it('is flattened to plain text past 1 MB, with a note, instead of building a markdown table in memory', async () => {
    const html = bigExport();
    expect(html.length).toBeGreaterThan(1_000_000);
    const answer = await convert(html);
    expect(answer.contentType).toBe('text/plain');
    expect(answer.text).toContain('Line | Region | Units | Status\nrow 0 | North region | 0 | on track\n');
    expect(answer.text.endsWith('row 14999 | North region | 44997 | on track')).toBe(true);
    expect(answer.text).not.toContain('<td>');
    expect(answer.text).not.toContain('Export');
    expect(answer.note).toContain(`This page is ${(html.length / 1_000_000).toFixed(1)} MB, over the 1 MB cap for a markdown render`);
    expect(answer.size).toBe(new TextEncoder().encode(answer.text).byteLength);
  });

  it('still renders a page of exactly 1 MB as markdown', async () => {
    const head = '<html><body><p>Kept</p><!--';
    const tail = '--></body></html>';
    const page = `${head}${'x'.repeat(1_000_000 - head.length - tail.length)}${tail}`;
    expect(new TextEncoder().encode(page).byteLength).toBe(1_000_000);
    expect(await convert(page)).toMatchObject({ contentType: 'text/markdown', text: 'Kept' });
  });

  it('still renders a page under the cap as markdown, tables included', async () => {
    const answer = await convert('<html><body><h1>Q3</h1><table><tr><th>Region</th><th>Units</th></tr><tr><td>North</td><td>12</td></tr></table></body></html>');
    expect(answer).toMatchObject({ contentType: 'text/markdown', text: '# Q3\n\n| Region | Units |\n| --- | --- |\n| North | 12 |' });
    expect(answer).not.toHaveProperty('note');
  });

  it('replaces an embedded data: image with a placeholder, as a Word file does, unless --inline-images true', async () => {
    const page = `<html><body><p>Revenue</p><img src="${PNG}" alt="chart"></body></html>`;
    const placeholder = await convert(page);
    expect(placeholder.text).toBe('Revenue\n\n[image: chart]');
    const inline = await convert(page, { inlineImages: 'true' });
    expect(inline.text).toContain('![chart](data:image/png;base64,');
  });
});
