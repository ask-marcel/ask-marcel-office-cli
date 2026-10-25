import { describe, expect, it } from 'bun:test';
import TurndownService from 'turndown';
import { htmlToMarkdown } from './turndown-adapter.ts';

describe('htmlToMarkdown — convert Graph-returned HTML (Office docs, OneNote, Outlook bodies) into clean markdown', () => {
  it('merges bold runs Outlook splits across tags, so a quoted header reads **发件人:** and not **发件人****:**', () => {
    // Outlook/Word emit the label and its colon as SEPARATE bold runs; turndown
    // closes and reopens, and the `****` seam shows as stray asterisks in
    // renderers that do not merge it. Two adjacent runs are one bold run.
    const result = htmlToMarkdown('<b>发件人</b><b>:</b> Robin Chen');

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toBe('**发件人:** Robin Chen');
      expect(result.value).not.toContain('****');
    }
  });

  it('merges a bold run split by Word’s nested spans, and three-way splits collapse to one run', () => {
    const nested = htmlToMarkdown('<b><span>差出人</span></b><b><span>:</span></b> Robin');
    const threeWay = htmlToMarkdown('<b>a</b><b>b</b><b>c</b>');

    if (nested.ok) expect(nested.value).toBe('**差出人:** Robin');
    // Each `</b><b>` boundary is dropped from the HTML, so all three runs become
    // one `<b>abc</b>` before turndown — not a half-merged **ab****c**.
    if (threeWay.ok) expect(threeWay.value).toBe('**abc**');
  });

  it('leaves bold runs that are genuinely separated by a space alone', () => {
    const result = htmlToMarkdown('<b>one</b> <b>two</b>');

    // `** **` is not `****`; merging here would delete a real word boundary.
    if (result.ok) expect(result.value).toBe('**one** **two**');
  });

  // The bold-run merge works on the HTML source, not the markdown output. These
  // four cases are why: a blunt replaceAll('****','') on the output corrupts
  // every place turndown does NOT escape asterisks. Each once regressed.
  it('keeps four literal asterisks inside a code span, which is content and not a bold seam', () => {
    const result = htmlToMarkdown('<code>a****b</code>');

    if (result.ok) expect(result.value).toBe('`a****b`');
  });

  it('keeps four literal asterisks inside a fenced code block', () => {
    const result = htmlToMarkdown('<pre><code>x = "****";</code></pre>');

    if (result.ok) expect(result.value).toContain('****');
  });

  it('keeps four asterisks that are part of a URL, so the link still resolves', () => {
    const result = htmlToMarkdown('<a href="https://contoso.example/a****b">link</a>');

    if (result.ok) expect(result.value).toBe('[link](https://contoso.example/a****b)');
  });

  it('does not blank out nested bold, whose delimiters are content the caller wrote', () => {
    // `<strong><b>x</b></strong>` is not two adjacent runs; deleting its `****`
    // would erase the word. Merging only same-tag ADJACENT runs leaves it be.
    const result = htmlToMarkdown('<strong><b>Total</b></strong>');

    if (result.ok) expect(result.value).toContain('Total');
  });

  it('does not splice a mixed b/strong pair into malformed markup', () => {
    // Only same-tag adjacency merges (`\1` backreference); a `</b><strong>`
    // boundary is left for turndown rather than stripped into `<b>…</strong>`.
    const result = htmlToMarkdown('<b>a</b><strong>b</strong>');

    if (result.ok) expect(result.value).toBe('**a****b**');
  });

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

describe('htmlToMarkdown — headerless (Outlook/Excel MSO) table degradation to pipe tables', () => {
  it('converts a pasted-from-Excel MsoNormalTable (all <td>, inline styles) to a pipe table instead of leaking raw HTML markup', () => {
    const html =
      '<table class="MsoNormalTable" border="0" cellspacing="0" style="border-collapse:collapse;mso-yfti-tbllook:1184">' +
      '<tr style="mso-yfti-irow:0"><td width="200" style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal">Name</p></td>' +
      '<td style="padding:0cm 5.4pt"><p class="MsoNormal">Amount</p></td></tr>' +
      '<tr><td style="padding:0cm 5.4pt"><p class="MsoNormal">Contoso A2</p></td>' +
      '<td style="padding:0cm 5.4pt"><p class="MsoNormal">1 200 <b>EUR</b></p></td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| Name | Amount |');
      expect(result.value).toContain('| --- | --- |');
      expect(result.value).toContain('| Contoso A2 | 1 200 **EUR** |');
      expect(result.value).not.toContain('<table');
      expect(result.value).not.toContain('MsoNormalTable');
    }
  });

  it('leaves a table with a proper <th> heading row on the GFM plugin path (unchanged behaviour)', () => {
    const html = '<table><thead><tr><th>H1</th><th>H2</th></tr></thead><tbody><tr><td>a</td><td>b</td></tr></tbody></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| H1 | H2 |');
      expect(result.value).toContain('| a | b |');
      expect(result.value).not.toContain('<table');
    }
  });

  it('unwraps a single-cell Outlook layout table and converts the inner data grid it wraps (the layout-outer / data-inner pattern)', () => {
    const inner = '<table><tr><td>k1</td><td>v1</td></tr><tr><td>k2</td><td>v2</td></tr></table>';
    const html = `<table width="600"><tr><td><p>Intro para</p>${inner}</td></tr></table>`;
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('Intro para\n\n| k1 | v1 |');
      expect(result.value).toContain('| k2 | v2 |');
      expect(result.value).not.toContain('<table');
    }
  });

  it('unwraps a single-column scaffold table into free-standing blocks instead of a one-column pipe table', () => {
    const html = '<table><tr><td><p>First block</p></td></tr><tr><td><p>Second block</p></td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('First block\n\nSecond block');
      expect(result.value).not.toContain('|');
    }
  });

  it('pads a colspan cell so later full-width rows still align', () => {
    const html = '<table><tr><td colspan="2">span</td></tr><tr><td>a</td><td>b</td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| span |  |');
      expect(result.value).toContain('| a | b |');
    }
  });

  it('escapes pipe characters inside cell text so the pipe table stays parseable', () => {
    const html = '<table><tr><td>a|b</td><td>c</td></tr><tr><td>d</td><td>e</td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toContain('| a\\|b | c |');
  });

  it('collapses a multi-paragraph cell onto one pipe-table line (pipe cells cannot hold newlines)', () => {
    const html = '<table><tr><td><p>line one</p><p>line two</p></td><td>x</td></tr><tr><td>a</td><td>b</td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| line one line two | x |');
      expect(result.value).toContain('| a | b |');
    }
  });

  it('flattens a nested table inside a data-grid cell to its text (a pipe table cannot nest another pipe table)', () => {
    const html = '<table><tr><td><table><tr><td>in1</td></tr><tr><td>in2</td></tr></table></td><td>x</td></tr><tr><td>a</td><td>b</td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| in1 in2 | x |');
      expect(result.value).not.toContain('<table');
    }
  });

  it('drops an empty <table></table> instead of crashing the GFM pass into tier-2 degradation', () => {
    const result = htmlToMarkdown('<p>x</p><table></table><p>y</p>');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('x');
      expect(result.value).toContain('y');
      expect(result.value).not.toContain('GFM table conversion failed');
      expect(result.value).not.toContain('<table');
    }
  });

  it('handles rows wrapped in an explicit <tbody> section (Outlook emits both bare and sectioned rows)', () => {
    const html = '<table><tbody><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></tbody></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| a | b |');
      expect(result.value).toContain('| c | d |');
      expect(result.value).not.toContain('<table');
    }
  });

  it('treats a non-numeric colspan attribute as a single column instead of throwing or padding', () => {
    const html = '<table><tr><td colspan="abc">a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| a | b |');
      expect(result.value).toContain('| c | d |');
    }
  });

  it('pads a ragged short row (fewer cells than the widest row) to the full column count', () => {
    const html = '<table><tr><td>a</td><td>b</td></tr><tr><td>only</td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| a | b |');
      expect(result.value).toContain('| only |  |');
    }
  });

  it('skips non-row table children (<caption>, <colgroup>) when building the degraded grid', () => {
    const html = '<table><caption>Budget</caption><colgroup><col><col></colgroup><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>';
    const result = htmlToMarkdown(html);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('| a | b |');
      expect(result.value).toContain('| c | d |');
      expect(result.value).not.toContain('<caption');
    }
  });

  it('drops a layout table whose only cell is empty (no blocks to unwrap)', () => {
    const result = htmlToMarkdown('<p>x</p><table><tr><td></td></tr></table><p>y</p>');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('x');
      expect(result.value).toContain('y');
      expect(result.value).not.toContain('|');
    }
  });
});
