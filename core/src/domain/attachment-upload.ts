/*
 * A file attached to a mail draft or to a To Do task. Graph takes a file under
 * 3 MB in one POST and a larger one through an upload session, up to a limit
 * per target (attachment-createuploadsession: 150 MB for Outlook, though the
 * tenant's message limit is 35 MB by default; taskfileattachment-
 * createuploadsession: 25 MB for To Do). The limits are decimal megabytes, the
 * safe reading of the documented "MB".
 */

/** Where a file is attached: an unsent mail draft, or a To Do task. */
type AttachmentTarget = { readonly kind: 'mail-draft'; readonly messageId: string } | { readonly kind: 'todo-task'; readonly listId: string; readonly taskId: string };

/** The file to attach: the name it shows under, its bytes, and its MIME type when known. */
type AttachmentFile = { readonly name: string; readonly bytes: Uint8Array; readonly contentType?: string };

/** The attachment made. `id` is absent only when Graph completed an upload session with no Location. */
type UploadedAttachment = { readonly id?: string; readonly name: string; readonly size: number };

/** The largest file Graph attaches to each target, in bytes. */
const ATTACHMENT_MAX_BYTES: Readonly<Record<AttachmentTarget['kind'], number>> = { 'mail-draft': 150_000_000, 'todo-task': 25_000_000 };

/** A file under this size goes in one POST; a file of this size or more, through an upload session. */
const SMALL_ATTACHMENT_BYTES = 3_000_000;

export { ATTACHMENT_MAX_BYTES, SMALL_ATTACHMENT_BYTES };
export type { AttachmentFile, AttachmentTarget, UploadedAttachment };
