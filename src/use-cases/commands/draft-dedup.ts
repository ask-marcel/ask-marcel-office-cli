// Client-side draft de-duplication. A conversationId $filter on the Drafts
// folder is not a reliable "is there already a draft on this thread" check
// (reply/forward drafts split across conversationIds, and $filter on Drafts is
// not read-your-writes consistent), so matching happens here on the normalized
// subject plus, optionally, a shared recipient. Pure and IO-free by design.

// Reply/forward subject prefixes across the major Outlook locales, longest
// first so `FWD` wins over `FW` and `RV` over `R`. Each must be followed by an
// optional count marker (AW[2]:) and a colon, so a real word that merely starts
// with a prefix letter ("Report:") is never mistaken for one. Built from a
// string constant (like mail-quote-stripper's label patterns) to keep the
// alternation readable and out of the regex-literal complexity budget.
const REPLY_FORWARD_TOKENS = 'FWD|ODP|FW|AW|WG|TR|RV|SV|VS|RE|R|V';
const REPLY_FORWARD_PREFIX = new RegExp(`^(?:(?:${REPLY_FORWARD_TOKENS})(?:\\[\\d+\\]|\\(\\d+\\))?\\s*[:：]\\s*)+`, 'i');

const normalizeThreadSubject = (subject: string): string => {
  const withoutPrefix = subject.trim().replace(REPLY_FORWARD_PREFIX, '');
  return withoutPrefix.replace(/\s+/g, ' ').trim().toLowerCase();
};

type DraftRecipient = { readonly emailAddress?: { readonly address?: string } };

type DraftForMatch = {
  readonly subject?: string;
  readonly toRecipients?: ReadonlyArray<DraftRecipient>;
  readonly ccRecipients?: ReadonlyArray<DraftRecipient>;
};

type DraftMatchCriteria = {
  readonly subject: string;
  readonly recipients?: ReadonlyArray<string>;
};

const draftAddresses = (draft: DraftForMatch): ReadonlyArray<string> =>
  [...(draft.toRecipients ?? []), ...(draft.ccRecipients ?? [])].flatMap((recipient) => {
    const address = recipient.emailAddress?.address;
    return address === undefined ? [] : [address.toLowerCase()];
  });

const shareRecipient = (draft: DraftForMatch, wanted: ReadonlyArray<string>): boolean => {
  const present = new Set(draftAddresses(draft));
  return wanted.some((address) => present.has(address.toLowerCase()));
};

const matchExistingDrafts = <T extends DraftForMatch>(drafts: ReadonlyArray<T>, criteria: DraftMatchCriteria): ReadonlyArray<T> => {
  const wantedSubject = normalizeThreadSubject(criteria.subject);
  const wantedRecipients = criteria.recipients ?? [];
  return drafts.filter((draft) => {
    if (normalizeThreadSubject(draft.subject ?? '') !== wantedSubject) return false;
    return wantedRecipients.length === 0 || shareRecipient(draft, wantedRecipients);
  });
};

export { matchExistingDrafts, normalizeThreadSubject };
export type { DraftForMatch, DraftMatchCriteria };
