import { describe, expect, it } from 'bun:test';
import TurndownService from 'turndown';
import { htmlToMarkdown } from './turndown-adapter.ts';

describe('htmlToMarkdown — convert Graph-returned HTML (Office docs, OneNote, Outlook bodies) into clean markdown', () => {
  it('renders headings as ATX (# H1) not setext underline', () => {
    const result = htmlToMarkdown('<h1>Title</h1><h2>Subtitle</h2>');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('# Title');
      expect(result.value).toContain('## Subtitle');
    }
  });

  it('emits fenced code blocks rather than indented ones', () => {
    const result = htmlToMarkdown('<pre><code>const x = 1;</code></pre>');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('```');
      expect(result.value).toContain('const x = 1;');
    }
  });

  it('preserves inline links with their text', () => {
    const result = htmlToMarkdown('<p>See <a href="https://example.com/docs">our docs</a>.</p>');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toContain('[our docs](https://example.com/docs)');
  });

  it('preserves images as ![alt](src) with the data URI intact', () => {
    const result = htmlToMarkdown('<p><img src="data:image/png;base64,iVBORw0=" alt="diagram"></p>');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toContain('![diagram](data:image/png;base64,iVBORw0=)');
  });

  it('renders unordered lists with `-` markers', () => {
    const result = htmlToMarkdown('<ul><li>one</li><li>two</li></ul>');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('-   one');
      expect(result.value).toContain('-   two');
    }
  });

  it('strips MSO conditional comments that pollute Outlook HTML', () => {
    const result = htmlToMarkdown('<p>before<!--[if !mso]> noisy MSO bracket <![endif]-->after</p>');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('before');
      expect(result.value).toContain('after');
      expect(result.value).not.toContain('mso');
      expect(result.value).not.toContain('endif');
    }
  });

  it('strips ordinary HTML comments', () => {
    const result = htmlToMarkdown('<p>visible<!-- secret note --></p>');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('visible');
      expect(result.value).not.toContain('secret');
    }
  });

  it('drops <script> and <style> blocks (turndown defaults handle these)', () => {
    const result = htmlToMarkdown('<p>kept</p><script>alert(1)</script><style>p{color:red}</style>');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('kept');
      expect(result.value).not.toContain('alert');
      expect(result.value).not.toContain('color:red');
    }
  });

  it('returns markdown_conversion_failed instead of throwing when turndown blows up on a malformed DOM', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = (() => {
      throw new TypeError("Cannot read properties of undefined (reading 'parentNode')");
    }) as typeof original;
    try {
      const result = htmlToMarkdown('<p>anything</p>');
      expect(result.ok).toBe(false);
      if (!result.ok && result.error.type === 'api_error') {
        expect(result.error.status).toBe(500);
        expect(result.error.message).toContain('markdown conversion failed');
        expect(result.error.message).toContain('parentNode');
      }
    } finally {
      proto.turndown = original;
    }
  });
});
