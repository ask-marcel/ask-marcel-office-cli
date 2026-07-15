// The slim default `$select` shared by the three mailbox READ commands
// (`list-mail-messages`, `search-mail-messages`, `get-mail-message`) so an
// unflagged call stays ~3 KB/result instead of shipping full HTML bodies. Kept
// as ONE constant so the flagship mail commands cannot silently drift onto
// different token-cost philosophies (they previously each held a copy of this
// string). `conversationId` is included so a caller can group results into a
// thread — and hand that id to `list-conversation-messages` — without a second
// round-trip. A user-supplied `--select` overrides this list entirely.
export const MAIL_MESSAGE_DEFAULT_SELECT = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview,conversationId';
