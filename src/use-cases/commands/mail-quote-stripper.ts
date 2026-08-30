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
 *   - Outlook mobile / new Outlook reference container:
 *     `<div id="mail-editor-reference-message-container">`, CONFIRMED by a
 *     header block opening inside it (the id alone wraps authored text too)
 *   - Outlook classic separator:                         `<hr id="stopSpelling">`
 *   - Gmail quote container:                             `<div class="gmail_quote">` / `<blockquote class="gmail_quote">`
 *   - Outlook reply divider (Word renderer / Mac / mobile):
 *     `<div style="…border-top:solid #E1E1E1|#B5C4DF 1.0pt…">`, and the same
 *     two hues in new Outlook's `border-color:rgb(…)` longhand
 *     (the colour constants are locale-independent across Outlook UI languages;
 *     #E1E1E1 is the desktop Word renderer, #B5C4DF the Mac / mobile reply rule)
 *   - Outlook desktop bold header block without any container marker:
 *     a `<b>`/`<strong>`-wrapped From label (localized) CONFIRMED by a
 *     companion Sent/Date label within a bounded window, so a body that
 *     merely bolds the word "From:" never truncates real content.
 */

const QUOTE_BOUNDARIES: ReadonlyArray<RegExp> = [
  /<div[^>]*\bid="divRplyFwdMsg"/i,
  /<div[^>]*\bid="appendonsend"/i,
  /<hr[^>]*\bid="stopSpelling"/i,
  /<div[^>]*\bclass="[^"]*\bgmail_quote\b/i,
  /<blockquote[^>]*\bclass="[^"]*\bgmail_quote\b/i,
  /<div[^>]*border-top:\s*solid #(?:E1E1E1|B5C4DF)/i,
  // The same two rules in new Outlook's longhand: the colour moves to
  // `border-color` in rgb() with the side chosen by `border-style: solid none
  // none`, which neither hex spelling above can match. Only these two hues are
  // recognized here; Word's generic `windowtext` border is deliberately left
  // out, since it draws ordinary bordered paragraphs and tables too.
  /<div[^>]*border-color:\s*rgb\((?:181,\s*196,\s*223|225,\s*225,\s*225)\)/i,
];

// Localized Outlook header labels. From-labels anchor a candidate cut; a
// companion Sent/Date-label within CONFIRM_WINDOW_CHARS confirms it as a real
// reply header (two bold labels a few hundred bytes apart is the Outlook
// header-block signature in every locale). Longer alternatives come first so
// `Enviado el` wins over `Enviado`.
//
// The second label is a "sent" word in some clients and a "date" word in
// others, for the SAME locale — a Chinese client quoted with 发件人/日期 where
// another writes 发件人/发送时间, and that whole message level survived stripping
// (live-reported 2026-07-17). Both vocabularies belong here.
//
// `Data` and `Date` are ambiguous outside their locale (an English body may
// legitimately bold "Data:"), so they widen the false-positive surface. That is
// accepted: a cut still needs a bold From-label within 400 chars, and a wrong
// cut is visible (it leaves the strip marker) and reversible (`--keep-quoted
// true`) rather than silent.
const FROM_LABELS = 'From|发件人|寄件者|差出人|보낸 사람|De|Von|Da|Van';
const SENT_LABELS = 'Sent|Datum|Date|Data|Fecha|发送时间|寄件日期|日期|送信日時|日付|보낸 날짜|날짜|Envoyé|Gesendet|Inviato|Enviado el|Enviado em|Enviado|Verzonden';
const CONFIRM_WINDOW_CHARS = 400;

// `(?:<span[^>]*>)?` tolerates the single MSO span Word nests inside the bold
// tag (`<b><span style='…'>From:</span></b>`); the leading `(?:\s|&nbsp;)*`
// covers the French "De :" space-before-colon and non-breaking-space variants;
// `[:：]` covers the CJK full-width colon.
//
// LABEL_COLON_GAP covers a Chinese Outlook web client that styles the colon
// differently from the label word and so emits them as two separate bold runs
// (`发件人</span></b><b><span lang=EN-HK>:`). Between the label and its colon we
// tolerate any run of inline open/close tags, not only whitespace. Every
// alternative is a tag or whitespace token, so the run still stops at the first
// character of real text: a bolded word followed by prose never reaches a later
// colon, keeping the From + (Sent|Date) false-positive guard intact.
//
// `<(?:b|strong)[^>]*>` also matched `<br>`, `<body>` and `<blockquote>`, so the
// "must be bold" half of the guard was never a bold check at all: a line break
// in front of a De:/Da:/Van: word, with any Date label in the next 400
// characters, truncated a body that quoted nothing. The lookahead requires the
// tag NAME to end where the match does (found 2026-08-30).
const BOLD_TAG = '<(?:b|strong)(?=[\\s>])[^>]*>';
const LABEL_COLON_GAP = `(?:\\s|&nbsp;|</span>|</b>|</strong>|<span[^>]*>|${BOLD_TAG})*`;
const HTML_FROM_LABEL = new RegExp(`${BOLD_TAG}(?:<span[^>]*>)?(?:\\s|&nbsp;)*(?:${FROM_LABELS})${LABEL_COLON_GAP}[:：]`, 'gi');
const HTML_SENT_LABEL = new RegExp(`${BOLD_TAG}(?:<span[^>]*>)?(?:\\s|&nbsp;)*(?:${SENT_LABELS})${LABEL_COLON_GAP}[:：]`, 'i');

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

// Outlook mobile and new Outlook wrap a quoted reply in
// `mail-editor-reference-message-container`, and the id was trusted on its own.
// It is not evidence of a quote: a message composed on Outlook mobile carried
// the author's ENTIRE body inside that container with nothing quoted below it,
// and the rule cut a 5,367-character message down to its 9-character greeting
// (live-reported 2026-08-30). Across all 16 messages of that thread, every
// container wrapping a genuine quote opened a confirmed From/Sent header within
// 229 characters of the container tag, while the false positive carried no
// header anywhere in the document — so the container cuts only when a header
// follows it, and an unrecognized locale now costs tokens rather than content.
//
// Scanned for every container, not just the first: an author who writes a long
// message inside the container still quotes below it, and the quote is what we
// want to cut at.
const MOBILE_CONTAINER = /<div[^>]*\bid="mail-editor-reference-message-container"[^>]*>/gi;
const CONTAINER_CONFIRM_WINDOW_CHARS = 600;

const confirmedContainerIndex = (html: string): number => {
  MOBILE_CONTAINER.lastIndex = 0;
  for (let match = MOBILE_CONTAINER.exec(html); match !== null; match = MOBILE_CONTAINER.exec(html)) {
    const windowStart = match.index + match[0].length;
    // The scan runs one confirmation window PAST the search window so a header
    // opening at the very end of it can still see its own Sent label.
    const header = confirmedHeaderIndex(html.slice(windowStart, windowStart + CONTAINER_CONFIRM_WINDOW_CHARS + CONFIRM_WINDOW_CHARS), HTML_FROM_LABEL, HTML_SENT_LABEL);
    if (header !== -1 && header < CONTAINER_CONFIRM_WINDOW_CHARS) return match.index;
  }
  return -1;
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
  const containerCut = confirmedContainerIndex(html);
  if (containerCut !== -1 && (cut === -1 || containerCut < cut)) cut = containerCut;
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
