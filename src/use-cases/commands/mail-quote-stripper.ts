import { z } from 'zod';
import type { CommandOptionMeta } from './command-types.ts';

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
 *   - Outlook desktop (Word renderer) separator:         `<div style="…border-top:solid #E1E1E1 1.0pt…">`
 *     (the hex constant is locale-independent across Outlook UI languages)
 *   - Outlook desktop bold header block without any container marker:
 *     a `<b>`/`<strong>`-wrapped From label (localized) CONFIRMED by a
 *     companion Sent/Date label within a bounded window, so a body that
 *     merely bolds the word "From:" never truncates real content.
 */

const QUOTE_BOUNDARIES: ReadonlyArray<RegExp> = [
  /<div[^>]*\bid="divRplyFwdMsg"/i,
  /<div[^>]*\bid="appendonsend"/i,
  /<div[^>]*\bid="mail-editor-reference-message-container"/i,
  /<hr[^>]*\bid="stopSpelling"/i,
  /<div[^>]*\bclass="[^"]*\bgmail_quote\b/i,
  /<blockquote[^>]*\bclass="[^"]*\bgmail_quote\b/i,
  /<div[^>]*border-top:\s*solid #E1E1E1/i,
];

// Localized Outlook header labels. From-labels anchor a candidate cut; a
// companion Sent/Date-label within CONFIRM_WINDOW_CHARS confirms it as a real
// reply header (two bold labels a few hundred bytes apart is the Outlook
// header-block signature in every locale). Longer alternatives come first so
// `Enviado el` wins over `Enviado`.
const FROM_LABELS = 'From|发件人|寄件者|差出人|보낸 사람|De|Von|Da|Van';
const SENT_LABELS = 'Sent|Date|发送时间|寄件日期|送信日時|보낸 날짜|Envoyé|Gesendet|Inviato|Enviado el|Enviado em|Enviado|Verzonden';
const CONFIRM_WINDOW_CHARS = 400;

// `(?:<span[^>]*>)?` tolerates the single MSO span Word nests inside the bold
// tag (`<b><span style='…'>From:</span></b>`); `(?:\s|&nbsp;)*` covers the
// French "De :" space-before-colon and non-breaking-space variants; `[:：]`
// covers the CJK full-width colon.
const HTML_FROM_LABEL = new RegExp(`<(?:b|strong)[^>]*>(?:<span[^>]*>)?(?:\\s|&nbsp;)*(?:${FROM_LABELS})(?:\\s|&nbsp;)*[:：]`, 'gi');
const HTML_SENT_LABEL = new RegExp(`<(?:b|strong)[^>]*>(?:<span[^>]*>)?(?:\\s|&nbsp;)*(?:${SENT_LABELS})(?:\\s|&nbsp;)*[:：]`, 'i');

const TEXT_FROM_LABEL = new RegExp(`^(?:${FROM_LABELS})\\s?[:：]`, 'gim');
const TEXT_SENT_LABEL = new RegExp(`^(?:${SENT_LABELS})\\s?[:：]`, 'im');

// Earliest index of a From-label whose confirmation window holds a
// Sent/Date-label, or -1. Shared by the HTML and plain-text strippers via the
// matching regex pair.
const confirmedHeaderIndex = (body: string, fromLabel: RegExp, sentLabel: RegExp): number => {
  fromLabel.lastIndex = 0;
  for (let match = fromLabel.exec(body); match !== null; match = fromLabel.exec(body)) {
    const windowStart = match.index + match[0].length;
    if (sentLabel.test(body.slice(windowStart, windowStart + CONFIRM_WINDOW_CHARS))) return match.index;
  }
  return -1;
};

// A bold From-label sits INSIDE its paragraph (`<p class=MsoNormal><b>From:`),
// so cutting at the label would leave the enclosing block's opening tag
// dangling before the strip marker. Walk the cut back over any opening
// block tags that immediately precede it.
const OPEN_BLOCK_TAIL = /<(?:p|div)[^>]*>\s*$/i;

const widenToBlockStart = (html: string, index: number): number => {
  let cut = index;
  for (let match = OPEN_BLOCK_TAIL.exec(html.slice(0, cut)); match !== null; match = OPEN_BLOCK_TAIL.exec(html.slice(0, cut))) {
    cut = match.index;
  }
  return cut;
};

const STRIP_MARKER = '<p><em>[Quoted reply chain removed — pass --keep-quoted true to include it]</em></p>';

// Plain-text reply boundaries, for `body.contentType === 'text'` messages where
// the HTML markers above don't apply. Conservative, line-anchored: the Outlook
// "Original Message" / underscore-rule banners, the Gmail/Apple "On <date> …
// wrote:" attribution line, and the first `>`-quoted line.
const PLAINTEXT_BOUNDARIES: ReadonlyArray<RegExp> = [/^-{2,}\s*Original Message\s*-{2,}\s*$/im, /^_{5,}\s*$/m, /^On\b.+\bwrote:\s*$/im, /^>.*$/m];

const PLAINTEXT_MARKER = '[Quoted reply chain removed — pass --keep-quoted true to include it]';

/**
 * Index where the quoted history begins in an HTML body, or -1 when the body
 * quotes nothing. Earliest of the structural markers merged with the widened
 * confirmed-header index. Exported because the draft-comment splicer needs the
 * same cut to place a reply ABOVE the quote without disturbing it — the stripper
 * throws the tail away, the splicer keeps it, but both agree on where it starts.
 */
const findQuoteBoundary = (html: string): number => {
  let cut = -1;
  for (const boundary of QUOTE_BOUNDARIES) {
    const match = boundary.exec(html);
    if (match !== null && (cut === -1 || match.index < cut)) cut = match.index;
  }
  const headerCut = confirmedHeaderIndex(html, HTML_FROM_LABEL, HTML_SENT_LABEL);
  const widenedHeaderCut = headerCut === -1 ? -1 : widenToBlockStart(html, headerCut);
  if (widenedHeaderCut !== -1 && (cut === -1 || widenedHeaderCut < cut)) cut = widenedHeaderCut;
  return cut;
};

/** The plain-text counterpart of `findQuoteBoundary`, for `contentType === 'text'` bodies. */
const findPlainTextQuoteBoundary = (text: string): number => {
  let cut = -1;
  for (const boundary of PLAINTEXT_BOUNDARIES) {
    const match = boundary.exec(text);
    if (match !== null && (cut === -1 || match.index < cut)) cut = match.index;
  }
  const headerCut = confirmedHeaderIndex(text, TEXT_FROM_LABEL, TEXT_SENT_LABEL);
  if (headerCut !== -1 && (cut === -1 || headerCut < cut)) cut = headerCut;
  return cut;
};

const stripQuotedReplies = (html: string): { readonly html: string; readonly stripped: boolean } => {
  const cut = findQuoteBoundary(html);
  if (cut === -1) return { html, stripped: false };
  return { html: `${html.slice(0, cut)}${STRIP_MARKER}`, stripped: true };
};

const stripQuotedPlainText = (text: string): { readonly text: string; readonly stripped: boolean } => {
  const cut = findPlainTextQuoteBoundary(text);
  if (cut === -1) return { text, stripped: false };
  return { text: `${text.slice(0, cut)}${PLAINTEXT_MARKER}`, stripped: true };
};

// Shared `--keep-quoted` schema field + option meta for every command that can
// hand an Outlook `.msg` to the markdown dispatch (the six converters below and
// `convert-mail-to-markdown`'s own Graph-body path). Declared once so the six
// descriptions cannot drift apart, and so the strip marker's remedy is real on
// every command that can emit it.
const keepQuotedSchemaField = z.enum(['true', 'false']).optional();

const keepQuotedOption: CommandOptionMeta = {
  name: 'keep-quoted',
  key: 'keepQuoted',
  required: false,
  description:
    'Applies to Outlook `.msg` input only. The quoted reply chain / forwarded-message block is stripped by default (it duplicates history and inflates the context budget) and replaced with a single visible marker naming this flag, so nothing is removed silently. Pass `--keep-quoted true` to render the full body. The recognized markers are the same set `convert-mail-to-markdown` uses — see `ask-marcel-office docs convert-mail-to-markdown`.',
  argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
};

export { findPlainTextQuoteBoundary, findQuoteBoundary, keepQuotedOption, keepQuotedSchemaField, stripQuotedPlainText, stripQuotedReplies };
