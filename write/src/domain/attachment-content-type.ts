/*
 * The MIME type of a file that a mail draft gets as an attachment, from the
 * extension of its name (in any case). Outlook shows and opens the file by
 * this type. An extension that is not in the table gets the general type.
 */

const TYPES: ReadonlyMap<string, string> = new Map(
  Object.entries({
    pdf: 'application/pdf',
    txt: 'text/plain',
    csv: 'text/csv',
    md: 'text/markdown',
    html: 'text/html',
    json: 'application/json',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    zip: 'application/zip',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  })
);

/** The type of a file with no known extension. */
export const GENERAL_CONTENT_TYPE = 'application/octet-stream';

/** The MIME type for the file `name`, from its extension. */
export const contentTypeOf = (name: string): string => {
  const dot = name.lastIndexOf('.');
  return (dot < 0 ? undefined : TYPES.get(name.slice(dot + 1).toLowerCase())) ?? GENERAL_CONTENT_TYPE;
};
