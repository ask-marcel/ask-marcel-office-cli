import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractDocxMetadata } from './docx-metadata.ts';
import { DOCX_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { extractPptxMetadata } from './pptx-metadata.ts';
import { extensionOf } from './text-passthrough.ts';
import { extractXlsxMetadata } from './xlsx-metadata.ts';

/**
 * The comments of a Word, Excel or PowerPoint file as one list, whatever the
 * format keeps them in: who wrote each, when, where it sits (the commented text
 * in a document, the cell in a workbook, the slide in a deck), what it says,
 * and whom it @-mentions. Built on the readers `--include-metadata` uses.
 *
 * A mention is a known name (the file's people registry, and everyone who
 * commented) written after an `@`, which is how all three apps store one in the
 * comment text. Excel writes every threaded comment a second time as a legacy
 * note starting `[Threaded comment]`, for versions that cannot show threads;
 * those copies are dropped so each comment is listed once.
 */

type DocumentComment = { readonly author: string; readonly date?: string; readonly anchor?: string; readonly text: string; readonly mentions: ReadonlyArray<string> };
type UnmentionedComment = Omit<DocumentComment, 'mentions'>;
type CommentFormat = 'docx' | 'xlsx' | 'pptx';
type DocumentComments = { readonly format: CommentFormat; readonly count: number; readonly comments: ReadonlyArray<DocumentComment> };
type CommentReader = (bytes: Uint8Array) => Promise<Result<ReadonlyArray<DocumentComment>, GraphError>>;

const THREADED_COPY = '[Threaded comment]';

const mentionsIn = (text: string, names: ReadonlyArray<string>): ReadonlyArray<string> => {
  const lower = text.toLowerCase();
  return [...new Set(names.filter((name) => name !== '' && lower.includes(`@${name.toLowerCase()}`)))];
};

// A sheet name that is not a plain word is quoted in a cell reference, its own quotes doubled.
const cellRef = (sheet: string | undefined, cell: string): string => {
  if (sheet === undefined) return cell;
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? `${sheet}!${cell}` : `'${sheet.replaceAll("'", "''")}'!${cell}`;
};

// The deck reader names a slide by its part (`slide3.xml`).
const slideAnchor = (slide: string | undefined): string | undefined => {
  const n = /^slide(\d+)\.xml$/.exec(slide ?? '')?.[1];
  return n === undefined ? slide : `slide ${n}`;
};

const withMentions = (comments: ReadonlyArray<UnmentionedComment>, people: ReadonlyArray<string>): ReadonlyArray<DocumentComment> => {
  const names = [...people, ...comments.map((c) => c.author)];
  return comments.map((c) => ({ ...c, mentions: mentionsIn(c.text, names) }));
};

const docxComments: CommentReader = async (bytes) => {
  const m = await extractDocxMetadata(bytes);
  if (!m.ok) return m;
  const comments = m.value.comments.map((c) => ({ author: c.author, date: c.date, anchor: c.anchor, text: c.text }));
  return ok(
    withMentions(
      comments,
      m.value.people.map((p) => p.author)
    )
  );
};

const xlsxComments: CommentReader = async (bytes) => {
  const m = await extractXlsxMetadata(bytes);
  if (!m.ok) return m;
  const notes = m.value.comments.filter((c) => !c.text.startsWith(THREADED_COPY)).map((c) => ({ author: c.author, anchor: cellRef(c.sheet, c.cell), text: c.text }));
  const threaded = m.value.threadedComments.map((c) => ({ author: c.author, date: c.date, anchor: cellRef(c.sheet, c.cell), text: c.text }));
  return ok(
    withMentions(
      [...notes, ...threaded],
      m.value.people.map((p) => p.displayName)
    )
  );
};

const pptxComments: CommentReader = async (bytes) => {
  const m = await extractPptxMetadata(bytes);
  if (!m.ok) return m;
  const comments = m.value.comments.map((c) => ({ author: c.author, date: c.date, anchor: slideAnchor(c.slide), text: c.text }));
  return ok(
    withMentions(
      comments,
      m.value.commentAuthors.map((a) => a.name)
    )
  );
};

const READERS: ReadonlyArray<readonly [ReadonlySet<string>, CommentFormat, CommentReader]> = [
  [DOCX_FAMILY, 'docx', docxComments],
  [XLSX_FAMILY, 'xlsx', xlsxComments],
  [PPTX_FAMILY, 'pptx', pptxComments],
];

const commentsFromBytes = async (bytes: Uint8Array, filename: string): Promise<Result<DocumentComments, GraphError>> => {
  const ext = extensionOf(filename);
  const reader = READERS.find(([family]) => family.has(ext));
  if (reader === undefined) {
    const what = ext === '' ? 'has no extension' : `is a .${ext}`;
    return err({
      type: 'api_error',
      status: 415,
      code: 'unsupported_format',
      message: `list-document-comments reads Word, Excel and PowerPoint files (docx, xlsx, pptx and their macro-enabled and template variants); this file ${what}.`,
    });
  }
  const [, format, read] = reader;
  const comments = await read(bytes);
  if (!comments.ok) return comments;
  return ok({ format, count: comments.value.length, comments: comments.value });
};

export { cellRef, commentsFromBytes, mentionsIn, slideAnchor };
export type { DocumentComment, DocumentComments };
