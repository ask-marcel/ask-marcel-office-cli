import { describe, expect, it } from 'bun:test';
import { buildMalformedDocx, buildMinimalOdt, buildRichOdp, buildRichOds, buildRichOdt } from '../../test-helpers/office-fixtures.ts';
import { odfContentToMarkdown, renderOdfContent } from './odf-content-to-markdown.ts';

const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"';
const ODF_NS_FULL =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"';

const text = (inner: string): string =>
  `<?xml version="1.0"?><office:document-content ${ODF_NS_FULL}><office:body><office:text>${inner}</office:text></office:body></office:document-content>`;

describe('odfContentToMarkdown — text document (.odt)', () => {
  it('renders headings (with level), inline runs (span/spaces/tab), nested lists, tables, and a display:none section in document order', async () => {
    const result = await odfContentToMarkdown(await buildRichOdt());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toBe(
      [
        '# Heading One',
        'First paragraph body.',
        '## Sub heading',
        'Spaced  out cell.',
        '- Item one\n- Item two\n  - Nested item',
        '| A1 | B1 |\n| --- | --- |\n| A2 | B2 |',
        '> _(hidden section — not shown in the rendered document)_',
        'Hidden body text',
        'Final paragraph.',
        '> 💬 **Reviewer, 2026-06-01T09:00:00**: Check this number.',
      ].join('\n\n')
    );
  });
});

describe('odfContentToMarkdown — office:annotation comments rendered inline', () => {
  it('renders an annotation as a blockquote line right after its paragraph, and does NOT leak the comment into the body text', () => {
    const md = renderOdfContent(
      text('<text:p>The body text<office:annotation><dc:creator>Alice</dc:creator><dc:date>2026-06-01T09:00:00</dc:date><text:p>my comment</text:p></office:annotation></text:p>')
    );
    expect(md).toBe('The body text\n\n> 💬 **Alice, 2026-06-01T09:00:00**: my comment');
  });

  it('falls back to just the comment when the annotation has no creator/date', () => {
    const md = renderOdfContent(text('<text:p>Body<office:annotation><text:p>orphan note</text:p></office:annotation></text:p>'));
    expect(md).toBe('Body\n\n> 💬 orphan note');
  });

  it('uses just the creator when there is no date', () => {
    const md = renderOdfContent(text('<text:p>Body<office:annotation><dc:creator>Bob</dc:creator><text:p>note</text:p></office:annotation></text:p>'));
    expect(md).toBe('Body\n\n> 💬 **Bob**: note');
  });

  it('finds an annotation nested inside an inline span', () => {
    const md = renderOdfContent(
      text('<text:p>Body <text:span>span<office:annotation><dc:creator>Carol</dc:creator><text:p>deep note</text:p></office:annotation></text:span></text:p>')
    );
    expect(md).toContain('> 💬 **Carol**: deep note');
    expect(md).toContain('Body span');
    expect(md).not.toContain('Body spandeep note'); // comment is not merged into the run text
  });

  it('emits only the annotation line for a paragraph that is just a comment', () => {
    const md = renderOdfContent(text('<text:p><office:annotation><dc:creator>Dan</dc:creator><text:p>standalone</text:p></office:annotation></text:p>'));
    expect(md).toBe('> 💬 **Dan**: standalone');
  });

  it('trims creator/date, drops empty body paragraphs, and joins multi-paragraph comment bodies with a single space', () => {
    const md = renderOdfContent(
      text(
        '<text:p>Body<office:annotation><dc:creator> Alice </dc:creator><dc:date> 2026 </dc:date><text:p> line one </text:p><text:p></text:p><text:p>line two</text:p></office:annotation></text:p>'
      )
    );
    // creator/date trimmed; the empty <text:p> filtered out; the two real lines joined by one space
    expect(md).toBe('Body\n\n> 💬 **Alice, 2026**: line one line two');
  });

  it('does not treat office:annotation-end as body content', () => {
    const md = renderOdfContent(text('<text:p>Body<office:annotation-end office:name="x"/> tail</text:p>'));
    expect(md).toBe('Body tail');
  });
});

describe('odfContentToMarkdown — spreadsheet (.ods)', () => {
  it('renders one named-heading + markdown table per sheet, capping and trimming the 16384-wide empty-cell tail', async () => {
    const result = await odfContentToMarkdown(await buildRichOds());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toBe(['## Budget', '| Item | Cost |\n| --- | --- |\n| Rent | 1000 |', '## Notes', '| Hello |\n| --- |'].join('\n\n'));
  });
});

describe('odfContentToMarkdown — presentation (.odp)', () => {
  it('renders one named-heading per slide followed by the slide text-box paragraphs', async () => {
    const result = await odfContentToMarkdown(await buildRichOdp());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toBe(['## Intro', 'Welcome slide', 'Subtitle here', '## Details', 'Detail bullet'].join('\n\n'));
  });
});

describe('odfContentToMarkdown — edges', () => {
  it('returns an empty string for a package with no content.xml', async () => {
    const result = await odfContentToMarkdown(await buildMinimalOdt());
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('');
  });

  it('propagates the zip-parse error for a malformed package', async () => {
    const result = await odfContentToMarkdown(buildMalformedDocx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('ooxml zip parse failed');
  });

  it('renderOdfContent returns empty string for undefined, empty, and body-less content', () => {
    expect(renderOdfContent(undefined)).toBe('');
    expect(renderOdfContent('')).toBe('');
    expect(renderOdfContent(`<?xml version="1.0"?><office:document-content ${ODF_NS}></office:document-content>`)).toBe('');
  });

  it('trims headings/paragraphs, drops empty paragraphs, applies text:s counts (default 1 when not > 0), and renders list-headers + nested defaults', () => {
    const md = renderOdfContent(
      text(
        '<text:h text:outline-level="3">  Padded Heading  </text:h>' +
          '<text:p>  </text:p>' +
          '<text:p>Word<text:s text:c="0"/>joined.</text:p>' +
          '<text:p>A<text:s text:c="3"/>B</text:p>' +
          '<text:list></text:list>' +
          '<text:list><text:list-item><text:p>  </text:p></text:list-item></text:list>' +
          '<text:list><text:list-header><text:p>Header item</text:p></text:list-header><text:list-item><text:p>Real item</text:p></text:list-item></text:list>' +
          '<table:table table:name="T"></table:table>' +
          '<text:section text:display="none"><text:p>  </text:p></text:section>' +
          '<text:section><text:p>Visible section text</text:p></text:section>' +
          '<text:h>Default level</text:h>'
      )
    );
    expect(md).toBe(['### Padded Heading', 'Word joined.', 'A   B', '- Header item\n- Real item', 'Visible section text', '# Default level'].join('\n\n'));
  });

  it('reads text:h and covered-table-cell content into cells and trims the trailing empty cell', () => {
    const md = renderOdfContent(
      text(
        '<table:table table:name="T2"><table:table-row>' +
          '<table:table-cell><text:h>HCell</text:h></table:table-cell>' +
          '<table:covered-table-cell><text:p>Covered</text:p></table:covered-table-cell>' +
          '<table:table-cell></table:table-cell>' +
          '</table:table-row></table:table>'
      )
    );
    expect(md).toBe('| HCell | Covered |\n| --- | --- |');
  });

  it('renders an unnamed sheet as "## Sheet" and an unnamed slide as "## Slide 1"', () => {
    const sheet = `<?xml version="1.0"?><office:document-content ${ODF_NS_FULL}><office:body><office:spreadsheet><table:table><table:table-row><table:table-cell><text:p>x</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    expect(renderOdfContent(sheet)).toBe('## Sheet\n\n| x |\n| --- |');
    const slides = `<?xml version="1.0"?><office:document-content ${ODF_NS_FULL}><office:body><office:presentation><draw:page><draw:frame><draw:text-box><text:p>Only slide</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>`;
    expect(renderOdfContent(slides)).toBe('## Slide 1\n\nOnly slide');
  });

  it('returns empty string for a body whose only child is whitespace text', () => {
    expect(renderOdfContent(`<?xml version="1.0"?><office:document-content ${ODF_NS}><office:body>   </office:body></office:document-content>`)).toBe('');
  });

  it('renders a text:line-break inside a paragraph as a space', () => {
    expect(renderOdfContent(text('<text:p>One<text:line-break/>Two</text:p>'))).toBe('One Two');
  });

  it('skips table:table-column, reads multi-paragraph cells (trimmed, empties dropped, joined by space)', () => {
    const md = renderOdfContent(
      text(
        '<table:table table:name="C"><table:table-column/><table:table-row><table:table-cell>' +
          '<text:p>  pad  </text:p><text:p></text:p><text:p>second</text:p>' +
          '</table:table-cell></table:table-row></table:table>'
      )
    );
    expect(md).toBe('| pad second |\n| --- |');
  });

  it('keeps a trailing row with a leading-empty cell but drops a fully-empty trailing row', () => {
    const md = renderOdfContent(
      text(
        '<table:table table:name="D">' +
          '<table:table-row><table:table-cell><text:p>A</text:p></table:table-cell><table:table-cell><text:p>B</text:p></table:table-cell></table:table-row>' +
          '<table:table-row><table:table-cell></table:table-cell><table:table-cell><text:p>C</text:p></table:table-cell></table:table-row>' +
          '<table:table-row><table:table-cell></table:table-cell><table:table-cell></table:table-cell></table:table-row>' +
          '</table:table>'
      )
    );
    expect(md).toBe('| A | B |\n| --- | --- |\n|  | C |');
  });

  it('falls back to "## Sheet" when a sheet element has attributes but no table:name, and omits the table line for an empty sheet', () => {
    const named = `<?xml version="1.0"?><office:document-content ${ODF_NS_FULL}><office:body><office:spreadsheet><table:table table:style-name="ta1"><table:table-row><table:table-cell><text:p>z</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    expect(renderOdfContent(named)).toBe('## Sheet\n\n| z |\n| --- |');
    const empty = `<?xml version="1.0"?><office:document-content ${ODF_NS_FULL}><office:body><office:spreadsheet><table:table table:name="Empty"></table:table></office:spreadsheet></office:body></office:document-content>`;
    expect(renderOdfContent(empty)).toBe('## Empty');
  });

  it('skips a leading whitespace text node when picking the body root element', () => {
    expect(
      renderOdfContent(
        `<?xml version="1.0"?><office:document-content ${ODF_NS_FULL}><office:body>  <office:text><text:p>X</text:p></office:text></office:body></office:document-content>`
      )
    ).toBe('X');
  });

  it('folds a table:table-header-rows group into the table body in order', () => {
    const xml =
      `<?xml version="1.0"?><office:document-content ${ODF_NS_FULL}><office:body><office:spreadsheet><table:table table:name="H">` +
      '<table:table-header-rows><table:table-row><table:table-cell><text:p>Col</text:p></table:table-cell></table:table-row></table:table-header-rows>' +
      '<table:table-row><table:table-cell><text:p>Val</text:p></table:table-cell></table:table-row>' +
      '</table:table></office:spreadsheet></office:body></office:document-content>';
    expect(renderOdfContent(xml)).toBe('## H\n\n| Col |\n| --- |\n| Val |');
  });
});
