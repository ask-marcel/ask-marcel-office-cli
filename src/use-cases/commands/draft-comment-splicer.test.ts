import { describe, expect, it } from 'bun:test';
import { escapeTextAsHtml, findBodyInsertStart, insertCommentAboveQuote, replaceCommentAboveQuote, replacePlainTextCommentAboveQuote } from './draft-comment-splicer.ts';

// The shape Graph actually returns from createReplyAll / createForward, split
// into the parts the splice has to respect. The <head> styles are load-bearing
// for the quoted tail's rendering, and the empty elementToProof div is where
// Graph parks the (empty) comment when the POST passes `comment: ''`.
const HEAD = '<html><head><style type="text/css" style="display:none"><!-- p {margin-top:0;margin-bottom:0} --></style></head>';
const BODY_OPEN = '<body dir="ltr">';
const EMPTY_COMMENT_DIV = '<div class="elementToProof" style="font-family:Aptos,Aptos_EmbeddedFont"></div>';
const QUOTE_TAIL =
  '<div id="appendonsend"></div><hr style="display:inline-block;width:98%" tabindex="-1">' +
  '<div id="divRplyFwdMsg" dir="ltr"><font face="Calibri"><b>From:</b> Robin Chen<br><b>Sent:</b> Monday, May 5, 2026<br><b>Subject:</b> Q3 planning</font></div>' +
  '<div>the original message body</div></body></html>';
const GRAPH_REPLY_DRAFT = `${HEAD}${BODY_OPEN}${EMPTY_COMMENT_DIV}${QUOTE_TAIL}`;

describe('placing an author’s reply text above the quoted history Graph minted', () => {
  it('inserts the reply above the quote while keeping the head styles, the body tag, and every byte of the quoted thread', () => {
    const result = insertCommentAboveQuote(GRAPH_REPLY_DRAFT, '<p>Confirmed for Concur.</p>');

    expect(result).toEqual({ html: `${HEAD}${BODY_OPEN}${EMPTY_COMMENT_DIV}<p>Confirmed for Concur.</p>${QUOTE_TAIL}`, boundaryFound: true });
  });

  it('replaces a previous reply with the revised one, dropping the div Graph minted and leaving the quoted thread whole', () => {
    // Revising twice must not stack: the second edit replaces the first reply
    // rather than inserting beside it, so the span from the body tag to the
    // quote is overwritten wholesale.
    const once = replaceCommentAboveQuote(GRAPH_REPLY_DRAFT, '<p>revised once</p>');
    const twice = replaceCommentAboveQuote(once.html, '<p>revised twice</p>');

    expect(once).toEqual({ html: `${HEAD}${BODY_OPEN}<p>revised once</p>${QUOTE_TAIL}`, boundaryFound: true });
    expect(twice).toEqual({ html: `${HEAD}${BODY_OPEN}<p>revised twice</p>${QUOTE_TAIL}`, boundaryFound: true });
  });

  it('leaves a draft with no quoted history untouched when revising, so the caller can refuse rather than overwrite the whole body', () => {
    const noQuote = `${HEAD}${BODY_OPEN}<div>a plain draft, nothing quoted</div></body></html>`;

    expect(replaceCommentAboveQuote(noQuote, '<p>revised</p>')).toEqual({ html: noQuote, boundaryFound: false });
  });

  it('falls back to just inside the body tag when creating on a draft that quotes nothing, keeping the head intact', () => {
    const noQuote = `${HEAD}${BODY_OPEN}<div>nothing quoted</div></body></html>`;

    expect(insertCommentAboveQuote(noQuote, '<p>new text</p>')).toEqual({
      html: `${HEAD}${BODY_OPEN}<p>new text</p><div>nothing quoted</div></body></html>`,
      boundaryFound: false,
    });
  });

  it('falls back to the very start when the draft body is a bare fragment with no body tag at all', () => {
    expect(insertCommentAboveQuote('<p>a fragment</p>', '<p>new text</p>')).toEqual({ html: '<p>new text</p><p>a fragment</p>', boundaryFound: false });
  });

  it('finds the insert point after the body tag whatever case the tag is written in, and at the start when there is none', () => {
    expect(findBodyInsertStart(GRAPH_REPLY_DRAFT)).toBe(`${HEAD}${BODY_OPEN}`.length);
    expect(findBodyInsertStart('<HTML><BODY LANG="EN">x</BODY></HTML>')).toBe('<HTML><BODY LANG="EN">'.length);
    expect(findBodyInsertStart('<p>a fragment</p>')).toBe(0);
  });

  it('escapes an author’s plain-text reply so markup they typed shows as characters and their line breaks survive as breaks', () => {
    const typed = 'Tom & Jerry <b>not bold</b> said "hi" to O\'Brien\r\nsecond line\nthird line';

    expect(escapeTextAsHtml(typed)).toBe('Tom &amp; Jerry &lt;b&gt;not bold&lt;/b&gt; said &quot;hi&quot; to O&#39;Brien<br>second line<br>third line');
  });

  it('escapes the ampersand before the angle brackets, so an entity the author typed is not double-escaped into nonsense', () => {
    expect(escapeTextAsHtml('&lt;')).toBe('&amp;lt;');
  });

  it('replaces the reply above a plain-text quote, separating it from the quote by a blank line', () => {
    const draft = 'the first attempt\n\n_______________________________\nFrom: Robin Chen\nSent: Monday, May 5, 2026\n\nthe original';

    expect(replacePlainTextCommentAboveQuote(draft, 'the revised reply')).toEqual({
      text: 'the revised reply\n\n_______________________________\nFrom: Robin Chen\nSent: Monday, May 5, 2026\n\nthe original',
      boundaryFound: true,
    });
  });

  it('leaves a plain-text draft with no quoted history untouched when revising', () => {
    expect(replacePlainTextCommentAboveQuote('nothing quoted here', 'revised')).toEqual({ text: 'nothing quoted here', boundaryFound: false });
  });
});
