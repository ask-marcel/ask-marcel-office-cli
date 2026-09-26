import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';
import type { MsgAttachment, MsgRecipient, MsgRecipientKind, ParsedMsg } from './msg-reader-adapter.ts';

/**
 * Raw RFC 822 messages (`.eml`) through postal-mime, mapped onto the `ParsedMsg`
 * shape the Outlook `.msg` reader produces, so both go through one renderer
 * (`renderMsg`): the same header block, quote stripping and recursive
 * attachments. An attached message stays an attachment (`forceRfc822Attachments`)
 * instead of being folded into the body text, and is named
 * `attached-message.eml` so the dispatch recurses into it the way it does into a
 * `.msg` inside a `.msg`. Dates print in the `.msg` reader's style
 * (`Mon, 14 Sep 2026 01:12:00 GMT`); a date the parser could not read is kept as
 * it came. `mapParsedEmail` is pure and exported so every branch is testable
 * without a crafted message.
 *
 * try/catch is permitted here per the infra-boundary rule: postal-mime throws
 * only past its own limits (a header block over 2 MB, MIME nesting past 256
 * levels), and that becomes a 415 like an unreadable `.msg`.
 */

type EmlMailbox = { readonly name: string; readonly address?: string; readonly group?: ReadonlyArray<EmlMailbox> };
type EmlAttachment = { readonly filename: string | null; readonly mimeType: string; readonly content: ArrayBuffer | Uint8Array | string };
type RawEml = {
  readonly subject?: string;
  readonly from?: EmlMailbox;
  readonly to?: ReadonlyArray<EmlMailbox>;
  readonly cc?: ReadonlyArray<EmlMailbox>;
  readonly bcc?: ReadonlyArray<EmlMailbox>;
  readonly date?: string;
  readonly text?: string;
  readonly html?: string;
  readonly attachments: ReadonlyArray<EmlAttachment>;
};

const nonEmpty = (value: string | undefined): string | undefined => (value === '' ? undefined : value);

// A group address (`Team: a@x, b@y;`) stands for its members.
const mailboxes = (list: ReadonlyArray<EmlMailbox>): ReadonlyArray<EmlMailbox> => list.flatMap((m) => m.group ?? [m]);

const recipients = (kind: MsgRecipientKind, list: ReadonlyArray<EmlMailbox> | undefined): ReadonlyArray<MsgRecipient> =>
  mailboxes(list ?? []).map((m) => ({ kind, name: nonEmpty(m.name), email: nonEmpty(m.address) }));

const attachmentOf = (a: EmlAttachment): MsgAttachment => ({
  fileName: a.filename ?? (a.mimeType === 'message/rfc822' ? 'attached-message.eml' : undefined),
  content: typeof a.content === 'string' ? new TextEncoder().encode(a.content) : new Uint8Array(a.content),
});

const msgStyleDate = (date: string | undefined): string | undefined => {
  const t = Date.parse(date ?? '');
  return Number.isNaN(t) ? date : new Date(t).toUTCString();
};

const mapParsedEmail = (raw: RawEml): ParsedMsg => {
  const sender = mailboxes(raw.from === undefined ? [] : [raw.from])[0];
  return {
    subject: raw.subject,
    senderName: nonEmpty(sender?.name),
    senderEmail: nonEmpty(sender?.address),
    date: msgStyleDate(raw.date),
    body: raw.text,
    bodyHtml: raw.html,
    recipients: [...recipients('to', raw.to), ...recipients('cc', raw.cc), ...recipients('bcc', raw.bcc)],
    attachments: raw.attachments.map(attachmentOf),
  };
};

const extractEml = async (bytes: Uint8Array): Promise<Result<ParsedMsg, GraphError>> => {
  try {
    const { default: PostalMime } = await import('postal-mime');
    return ok(mapParsedEmail(await PostalMime.parse(bytes, { forceRfc822Attachments: true })));
  } catch (e) {
    return err({ type: 'api_error', status: 415, message: `failed to parse .eml (RFC 822 message): ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { extractEml, mapParsedEmail };
