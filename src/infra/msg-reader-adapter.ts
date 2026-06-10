import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Parse an Outlook `.msg` file — the OLE Compound File (CFBF) container Outlook
 * writes when you drag an email to disk, same binary family as legacy .doc/.xls —
 * into a clean {@link ParsedMsg} via @kenjiuno/msgreader (pure-JS, no native deps,
 * lazy-imported like word-extractor / unpdf so it stays out of the cold-start path).
 *
 * The adapter is the anti-corruption layer: msgreader returns a wide `FieldsData`
 * (X500 sender addresses, RTF body, hundreds of MAPI props); {@link mapRawMsg}
 * narrows it to the half-dozen fields the markdown renderer needs and resolves the
 * sender SMTP / delivery-time fallbacks. `mapRawMsg` is pure and exported so every
 * fallback branch is unit-tested with plain objects — the binary fixture only
 * exercises one field combination.
 *
 * try/catch is permitted here per the infra-boundary rule: msgreader throws on a
 * non-OLE / corrupt container, and a single unreadable attachment must not sink the
 * whole parse, so attachment reads are individually guarded.
 */

type MsgRecipientKind = 'to' | 'cc' | 'bcc' | 'unknown';
type MsgRecipient = { readonly kind: MsgRecipientKind; readonly name?: string; readonly email?: string };
type MsgAttachment = { readonly fileName?: string; readonly content?: Uint8Array };
type ParsedMsg = {
  readonly subject?: string;
  readonly senderName?: string;
  readonly senderEmail?: string;
  readonly date?: string;
  readonly body?: string;
  readonly bodyHtml?: string;
  readonly recipients: readonly MsgRecipient[];
  readonly attachments: readonly MsgAttachment[];
};

// Structural view of the bits of msgreader's `FieldsData` we consume.
type RawRecipient = { readonly recipType?: string; readonly name?: string; readonly smtpAddress?: string; readonly email?: string };
type RawAttachment = { readonly fileName?: string };
type RawMsg = {
  readonly subject?: string;
  readonly senderName?: string;
  readonly senderSmtpAddress?: string;
  readonly senderEmail?: string;
  readonly messageDeliveryTime?: string;
  readonly clientSubmitTime?: string;
  readonly body?: string;
  readonly bodyHtml?: string;
  readonly recipients?: readonly RawRecipient[];
  readonly attachments?: readonly RawAttachment[];
};

const KNOWN_KINDS: ReadonlySet<string> = new Set(['to', 'cc', 'bcc']);

// msgreader already maps the MAPI PR_RECIPIENT_TYPE integer to 'to'/'cc'/'bcc';
// it is `undefined` only when the recipient row carries no type property.
const toKind = (value: string | undefined): MsgRecipientKind => (value !== undefined && KNOWN_KINDS.has(value) ? (value as MsgRecipientKind) : 'unknown');

const mapRawMsg = (raw: RawMsg, attachmentContents: readonly (Uint8Array | undefined)[]): ParsedMsg => ({
  subject: raw.subject,
  senderName: raw.senderName,
  // PR_SENDER_EMAIL_ADDRESS is the X500 legacyExchangeDN for internal senders;
  // prefer the resolved SMTP address, fall back to whatever is present.
  senderEmail: raw.senderSmtpAddress ?? raw.senderEmail,
  date: raw.messageDeliveryTime ?? raw.clientSubmitTime,
  body: raw.body,
  bodyHtml: raw.bodyHtml,
  recipients: (raw.recipients ?? []).map((r) => ({ kind: toKind(r.recipType), name: r.name, email: r.smtpAddress ?? r.email })),
  attachments: (raw.attachments ?? []).map((a, index) => ({ fileName: a.fileName, content: attachmentContents[index] })),
});

type AttachmentReader = { readonly getAttachment: (index: number) => { readonly content?: unknown } };

// msgreader is a CJS default-export class. The dynamic-import `default` resolves to a
// constructable type under the dev tsconfig but to the module namespace under the
// declaration-emit tsconfig (a CJS/ESM interop quirk). Pin the shape we use so both
// configs agree and the constructor is always callable.
type MsgReaderInstance = AttachmentReader & { readonly getFileData: () => unknown };
type MsgReaderCtor = new (input: ArrayBuffer | DataView) => MsgReaderInstance;

const readAttachmentContent = (reader: AttachmentReader, index: number): Uint8Array | undefined => {
  try {
    const content = reader.getAttachment(index).content;
    return content instanceof Uint8Array ? content : undefined;
  } catch {
    return undefined;
  }
};

/**
 * Resolve the MsgReader class from the dynamic import's `default`, whichever
 * interop shape it arrives in. msgreader is CJS with `exports.default = class`
 * (+ `__esModule`): Bun's RUNTIME import hands us the class directly, but Bun's
 * BUNDLER (node-mode `__toESM`) sets `default` to the whole exports object —
 * so in `dist/cli.js` the class sits one level deeper at `default.default`
 * ("Object is not a constructor" at runtime, invisible to source-run tests).
 */
const resolveMsgReaderCtor = (defaultExport: unknown): MsgReaderCtor => {
  if (typeof defaultExport === 'function') return defaultExport as unknown as MsgReaderCtor;
  return (defaultExport as { readonly default: MsgReaderCtor }).default;
};

const extractMsg = async (bytes: Uint8Array): Promise<Result<ParsedMsg, GraphError>> => {
  try {
    const MsgReader = resolveMsgReaderCtor((await import('@kenjiuno/msgreader')).default);
    // The constructor accepts `ArrayBuffer | DataView`; wrap the (possibly offset)
    // view zero-copy rather than copying the bytes into a fresh buffer.
    const reader = new MsgReader(new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength));
    const data = reader.getFileData() as RawMsg;
    const contents = (data.attachments ?? []).map((_attachment, index) => readAttachmentContent(reader, index));
    return ok(mapRawMsg(data, contents));
  } catch (e) {
    return err({ type: 'api_error', status: 415, message: `failed to parse .msg (Outlook message): ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { extractMsg, mapRawMsg, readAttachmentContent, resolveMsgReaderCtor };
export type { MsgAttachment, MsgRecipient, MsgRecipientKind, ParsedMsg };
