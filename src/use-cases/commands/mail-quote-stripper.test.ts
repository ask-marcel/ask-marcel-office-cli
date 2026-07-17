import { describe, expect, it } from 'bun:test';
import { findPlainTextQuoteBoundary, findQuoteBoundary, stripQuotedPlainText, stripQuotedReplies } from './mail-quote-stripper.ts';

// The markers the strippers append. Hardcoded rather than imported: they are the
// user-visible contract of `--keep-quoted`, so a test that reads them from the
// module under test would pass even if the wording silently changed.
const STRIP_MARKER = '<p><em>[Quoted reply chain removed — pass --keep-quoted true to include it]</em></p>';
const PLAINTEXT_MARKER = '[Quoted reply chain removed — pass --keep-quoted true to include it]';

describe('locating where a reply body stops being the author’s new text and starts quoting the thread', () => {
  it('points at the divider Outlook writes above the quote, not at the From/Sent header further down, in a real reply carrying both', () => {
    // The real Outlook shape: `appendonsend` (the "type above this line" marker)
    // precedes `divRplyFwdMsg`, which itself wraps a bold From:/Sent: pair. All
    // three markers fire; the earliest must win, or the reply keeps dead history.
    const authorText = '<p>Confirmed for Contoso.</p>';
    const quote = '<div id="appendonsend"></div><hr><div id="divRplyFwdMsg"><b>From:</b> Robin Chen<br><b>Sent:</b> Monday, May 5, 2026<br><b>To:</b> Alex Kim</div>';
    const html = `${authorText}${quote}<p>the original</p>`;

    expect(findQuoteBoundary(html)).toBe(authorText.length);
  });

  it('points at the earliest quote marker even when the one appearing first in the document is later in the boundary list', () => {
    // divRplyFwdMsg is 1st in QUOTE_BOUNDARIES, gmail_quote is 5th; here the
    // gmail one sits later in the BODY. A boundary scan that lets each match
    // overwrite the last would cut at the gmail div and keep dead quoted text.
    const authorText = '<p>keep</p>';
    const html = `${authorText}<div id="divRplyFwdMsg">A</div><p>mid</p><div class="gmail_quote">B</div>`;

    expect(findQuoteBoundary(html)).toBe(authorText.length);
  });

  it('reports no boundary for an email that quotes nothing, so a blockquote around a customer’s own sentence is never mistaken for a reply chain', () => {
    expect(findQuoteBoundary('<p>Agreed with <blockquote>ship it Friday</blockquote> as written.</p>')).toBe(-1);
  });

  it('points at the enclosing paragraph, not the bold label inside it, when Outlook desktop writes the From/Sent header pair with no container marker', () => {
    const authorText = '<p>Confirmed.</p>';
    const header = "<p class=MsoNormal><b><span style='font-size:11.0pt'>From:</span></b> Robin Chen<br><b>Sent:</b> Monday, May 5, 2026</p>";

    // Cutting at the `<b>` would leave `<p class=MsoNormal>` dangling before the
    // marker; the cut widens back over the block tag that opens the header.
    expect(findQuoteBoundary(`${authorText}${header}<p>the original</p>`)).toBe(authorText.length);
  });

  it('widens the cut back over every block tag opening a header, not just the innermost one, when Outlook nests the header pair in a div', () => {
    const authorText = '<p>Confirmed.</p>';
    const header = '<div><p class=MsoNormal><b>From:</b> Robin Chen<br><b>Sent:</b> Monday, May 5, 2026</p></div>';

    // Two block tags open before the bold label (`<div>` then `<p>`); walking back
    // over only the innermost would still strand the `<div>` before the marker.
    expect(findQuoteBoundary(`${authorText}${header}`)).toBe(authorText.length);
  });

  it('points at the underscore rule Outlook classic draws above a quoted plain-text original', () => {
    const authorText = 'Confirmed for Contoso.\n\n';

    expect(findPlainTextQuoteBoundary(`${authorText}_______________________________\nFrom: Robin Chen`)).toBe(authorText.length);
  });

  it('points at the attribution line above a quoted plain-text original, not at the >-quoted lines below it', () => {
    const authorText = 'Sounds good.\n\n';
    const text = `${authorText}On Mon, May 5, 2026 at 3:00 PM Robin Chen <robin.chen@contoso.com> wrote:\n> the original`;

    expect(findPlainTextQuoteBoundary(text)).toBe(authorText.length);
  });

  it('points at the underscore rule rather than the From/Sent header below it when a plain-text reply carries both', () => {
    const authorText = 'Confirmed.\n\n';
    const text = `${authorText}_______________________________\nFrom: Robin Chen\nSent: Monday, May 5, 2026\nTo: Alex Kim\n\nold body`;

    // The header pair is a real boundary too, but it sits BELOW the rule; cutting
    // there would leave the bare rule dangling at the end of the kept text.
    expect(findPlainTextQuoteBoundary(text)).toBe(authorText.length);
  });

  it('points at the Original Message banner Outlook draws above a quoted plain-text original', () => {
    const authorText = 'Reply.\n\n';

    expect(findPlainTextQuoteBoundary(`${authorText}-----Original Message-----\nFrom: Robin Chen`)).toBe(authorText.length);
  });

  it('reports no boundary when an Original Message banner is quoted mid-sentence rather than drawn on its own line', () => {
    // The banner is line-anchored at both ends: prose that merely mentions it, or
    // a line that continues past it, is the author's own text and must survive.
    expect(findPlainTextQuoteBoundary('See the -----Original Message----- banner below for the format.')).toBe(-1);
    expect(findPlainTextQuoteBoundary('Reply.\n-----Original Message----- and then some trailing prose.')).toBe(-1);
  });

  it('reports no boundary for a plain-text note that mentions From: without a companion Sent line, so ordinary prose is never truncated', () => {
    expect(findPlainTextQuoteBoundary('Note to self.\nFrom: the spec, all requests need auth.\nMore notes follow here.')).toBe(-1);
  });

  it('strips an HTML reply at exactly the boundary the finder reports, and leaves a quote-free body whole', () => {
    const html = '<p>keep</p><div id="divRplyFwdMsg">quoted</div>';
    const quoteFree = '<p>nothing quoted here</p>';

    expect(stripQuotedReplies(html)).toEqual({ html: `${html.slice(0, findQuoteBoundary(html))}${STRIP_MARKER}`, stripped: true });
    expect(stripQuotedReplies(quoteFree)).toEqual({ html: quoteFree, stripped: false });
  });

  it('strips a plain-text reply at exactly the boundary the finder reports, and leaves a quote-free body whole', () => {
    const text = 'keep\n\n_______________________________\nFrom: Robin Chen';
    const quoteFree = 'nothing quoted here';

    expect(stripQuotedPlainText(text)).toEqual({ text: `${text.slice(0, findPlainTextQuoteBoundary(text))}${PLAINTEXT_MARKER}`, stripped: true });
    expect(stripQuotedPlainText(quoteFree)).toEqual({ text: quoteFree, stripped: false });
  });
});
