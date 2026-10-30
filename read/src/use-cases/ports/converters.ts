import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';

/*
 * The document converters a read command runs on bytes it fetched (ADR 0004,
 * decision 5): HTML to markdown, the OOXML and plain zip readers, the PDF, DOC,
 * DOCX, MSG, EML and XLSX extractors, and the line diff. Infra builds the real
 * bag (src/infra/converters.ts) on the npm libraries; the composition root
 * passes it to the commands, so no use-case imports an adapter.
 */

type OoxmlZip = {
  readonly read: (path: string) => string | undefined;
  readonly list: () => ReadonlyArray<string>;
};

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

type MediaPart = { readonly path: string; readonly bytes: Uint8Array };
type ZipEntry = { readonly path: string; readonly bytes: Uint8Array };
type PdfText = { readonly text: string; readonly pageCount: number };
type SheetCsv = { readonly name: string; readonly csv: string };
type LineDiff = { readonly patch: string; readonly added: number; readonly removed: number };

type Converters = {
  readonly htmlToMarkdown: (html: string) => Result<string, GraphError>;
  readonly stripHtmlToText: (html: string) => string;
  readonly mammothToHtml: (bytes: Uint8Array) => Promise<Result<string, GraphError>>;
  readonly openOoxmlZip: (bytes: Uint8Array) => Promise<Result<OoxmlZip, GraphError>>;
  readonly openZipEntries: (bytes: Uint8Array) => Promise<Result<ReadonlyArray<ZipEntry>, GraphError>>;
  readonly extractOoxmlMedia: (bytes: Uint8Array) => Promise<Result<ReadonlyArray<MediaPart>, GraphError>>;
  readonly extractPdfImages: (bytes: Uint8Array, includePage?: (page: number) => boolean) => Promise<Result<ReadonlyArray<MediaPart>, GraphError>>;
  readonly extractPdfText: (bytes: Uint8Array) => Promise<Result<PdfText, GraphError>>;
  readonly extractDocText: (bytes: Uint8Array) => Promise<Result<string, GraphError>>;
  readonly extractMsg: (bytes: Uint8Array) => Promise<Result<ParsedMsg, GraphError>>;
  readonly extractEml: (bytes: Uint8Array) => Promise<Result<ParsedMsg, GraphError>>;
  readonly readSheetsAsCsv: (bytes: Uint8Array) => Promise<Result<ReadonlyArray<SheetCsv>, GraphError>>;
  readonly lineDiff: (fromName: string, fromText: string, toName: string, toText: string) => LineDiff | undefined;
};

export type { Converters, LineDiff, MediaPart, MsgAttachment, MsgRecipient, MsgRecipientKind, OoxmlZip, ParsedMsg, PdfText, SheetCsv, ZipEntry };
