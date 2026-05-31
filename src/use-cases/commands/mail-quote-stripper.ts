/**
 * Strips quoted reply chains / forwarded-message blocks from an Outlook or
 * Gmail HTML email body so long threads don't duplicate quoted content into the
 * model's context. Conservative: truncates the body at the EARLIEST well-known
 * reply/forward boundary marker and replaces the tail with a single visible
 * placeholder — nothing is removed silently, and `--keep-quoted true` on
 * `convert-mail-to-markdown` restores the full body. Pure string transform.
 *
 * Only structural, vendor-specific markers are matched (never a bare
 * `<blockquote>`, which legitimate content uses too):
 *   - Outlook desktop / OWA reply+forward header block: `<div id="divRplyFwdMsg">`
 *   - Outlook "type above this line" boundary:           `<div id="appendonsend">`
 *   - Outlook mobile reference container:                `<div id="mail-editor-reference-message-container">`
 *   - Outlook classic separator:                         `<hr id="stopSpelling">`
 *   - Gmail quote container:                             `<div class="gmail_quote">` / `<blockquote class="gmail_quote">`
 */

const QUOTE_BOUNDARIES: ReadonlyArray<RegExp> = [
  /<div[^>]*\bid="divRplyFwdMsg"/i,
  /<div[^>]*\bid="appendonsend"/i,
  /<div[^>]*\bid="mail-editor-reference-message-container"/i,
  /<hr[^>]*\bid="stopSpelling"/i,
  /<div[^>]*\bclass="[^"]*\bgmail_quote\b/i,
  /<blockquote[^>]*\bclass="[^"]*\bgmail_quote\b/i,
];

const STRIP_MARKER = '<p><em>[Quoted reply chain removed — pass --keep-quoted true to include it]</em></p>';

// Plain-text reply boundaries, for `body.contentType === 'text'` messages where
// the HTML markers above don't apply. Conservative, line-anchored: the Outlook
// "Original Message" / underscore-rule banners, the Gmail/Apple "On <date> …
// wrote:" attribution line, and the first `>`-quoted line.
const PLAINTEXT_BOUNDARIES: ReadonlyArray<RegExp> = [/^-{2,}\s*Original Message\s*-{2,}\s*$/im, /^_{5,}\s*$/m, /^On\b.+\bwrote:\s*$/im, /^>.*$/m];

const PLAINTEXT_MARKER = '[Quoted reply chain removed — pass --keep-quoted true to include it]';

const stripQuotedReplies = (html: string): { readonly html: string; readonly stripped: boolean } => {
  let cut = -1;
  for (const boundary of QUOTE_BOUNDARIES) {
    const match = boundary.exec(html);
    if (match !== null && (cut === -1 || match.index < cut)) cut = match.index;
  }
  if (cut === -1) return { html, stripped: false };
  return { html: `${html.slice(0, cut)}${STRIP_MARKER}`, stripped: true };
};

const stripQuotedPlainText = (text: string): { readonly text: string; readonly stripped: boolean } => {
  let cut = -1;
  for (const boundary of PLAINTEXT_BOUNDARIES) {
    const match = boundary.exec(text);
    if (match !== null && (cut === -1 || match.index < cut)) cut = match.index;
  }
  if (cut === -1) return { text, stripped: false };
  return { text: `${text.slice(0, cut)}${PLAINTEXT_MARKER}`, stripped: true };
};

export { stripQuotedPlainText, stripQuotedReplies };
