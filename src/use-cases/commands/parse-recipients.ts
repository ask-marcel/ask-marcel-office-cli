/**
 * Parse a comma-separated recipient string into Microsoft Graph address
 * objects. Shared by the mail-draft write commands (create-mail-draft,
 * update-mail-draft, create-forward-draft) so the split/trim/empty-drop
 * behaviour stays identical across all three. Whitespace around each address
 * is trimmed and empty segments (leading, trailing, or doubled commas) are
 * dropped, so `"a@x.com, , b@x.com,"` yields exactly two recipients.
 */
const parseRecipients = (csv: string): Array<{ emailAddress: { address: string } }> =>
  csv
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
    .map((address) => ({ emailAddress: { address } }));

export { parseRecipients };
