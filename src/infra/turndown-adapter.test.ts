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

  it('retries turndown WITHOUT GFM when the first pass throws, so most Outlook MSO bodies still convert to clean markdown (tier 2)', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    let callCount = 0;
    proto.turndown = function (this: TurndownService, input: string): string {
      callCount += 1;
      if (callCount === 1) throw new TypeError("Cannot read properties of undefined (reading 'parentNode')");
      return original.call(this, input);
    };
    try {
      const result = htmlToMarkdown('<p>Hello <b>world</b>.</p>');
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.value).toContain('GFM table conversion failed');
        expect(result.value).toContain('parentNode');
        expect(result.value).toContain('tables flattened to paragraphs');
        expect(result.value).toContain('Hello **world**.');
        expect(callCount).toBe(2);
      }
    } finally {
      proto.turndown = original;
    }
  });

  it('falls back to stripped-text body (with a markdown note prefix) when BOTH turndown passes throw, so the LLM still gets readable content (tier 3)', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = () => {
      throw new TypeError("Cannot read properties of undefined (reading 'parentNode')");
    };
    try {
      const result = htmlToMarkdown('<p>Hello <b>world</b>.</p><script>alert(1)</script>');
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.value).toContain('markdown conversion failed');
        expect(result.value).toContain('parentNode');
        expect(result.value).toContain('Hello world.');
        expect(result.value).not.toContain('alert(1)');
      }
    } finally {
      proto.turndown = original;
    }
  });

  it('decodes basic HTML entities in the stripped-text fallback path', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = () => {
      throw new TypeError('boom');
    };
    try {
      const result = htmlToMarkdown('<p>tom &amp; jerry &lt;3 &nbsp;always</p>');
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.value).toContain('tom & jerry <3 always');
    } finally {
      proto.turndown = original;
    }
  });

  it('emits only the failure note when the input HTML is empty (so callers do not get a blank fallback string)', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = () => {
      throw new TypeError('boom');
    };
    try {
      const result = htmlToMarkdown('');
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.value).toContain('markdown conversion failed');
        expect(result.value).not.toContain('\n\n');
      }
    } finally {
      proto.turndown = original;
    }
  });

  it('inserts a newline at <br>, <br/>, and </p> in the stripped-text fallback', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = () => {
      throw new TypeError('boom');
    };
    try {
      const result = htmlToMarkdown('<p>line1</p><p>line2</p>line3<br>line4<br/>line5');
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.value).toContain('line1');
        expect(result.value).toContain('line2');
        expect(result.value).toContain('line3');
        expect(result.value).toContain('line4');
        expect(result.value).toContain('line5');
        const body = result.value.split('\n\n').slice(1).join('\n\n');
        expect(body.split('\n').length).toBeGreaterThan(1);
      }
    } finally {
      proto.turndown = original;
    }
  });

  it('strips HTML comments from the stripped-text fallback (both well-formed and unclosed)', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = () => {
      throw new TypeError('boom');
    };
    try {
      const wellFormed = htmlToMarkdown('<p>before<!-- secret -->after</p>');
      const unclosed = htmlToMarkdown('<p>visible<!-- never closed');
      expect(wellFormed.ok).toBe(true);
      if (wellFormed.ok) {
        expect(wellFormed.value).toContain('before');
        expect(wellFormed.value).toContain('after');
        expect(wellFormed.value).not.toContain('secret');
      }
      expect(unclosed.ok).toBe(true);
      if (unclosed.ok) {
        expect(unclosed.value).toContain('visible');
        expect(unclosed.value).not.toContain('never closed');
      }
    } finally {
      proto.turndown = original;
    }
  });

  it('drops unclosed <script> bodies in the stripped-text fallback (no closing </script> in the input)', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = () => {
      throw new TypeError('boom');
    };
    try {
      const result = htmlToMarkdown('<p>kept</p><script>alert(1)');
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.value).toContain('kept');
        expect(result.value).not.toContain('alert(1)');
      }
    } finally {
      proto.turndown = original;
    }
  });

  it('handles HTML whose final fragment has no `<` (the no-more-tags branch)', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = () => {
      throw new TypeError('boom');
    };
    try {
      const result = htmlToMarkdown('<p>start</p>just trailing text with no more tags');
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.value).toContain('start');
        expect(result.value).toContain('just trailing text with no more tags');
      }
    } finally {
      proto.turndown = original;
    }
  });

  it('stops cleanly when an HTML tag is unclosed (no `>` found after the `<`)', () => {
    const proto = (TurndownService as unknown as { prototype: { turndown: (input: string) => string } }).prototype;
    const original = proto.turndown;
    proto.turndown = () => {
      throw new TypeError('boom');
    };
    try {
      const result = htmlToMarkdown('<p>before</p><span class="never-closed');
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.value).toContain('before');
    } finally {
      proto.turndown = original;
    }
  });
});
