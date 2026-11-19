/*
 * The subject and the body of a calendar draft. The subject starts with
 * `[DRAFT] `. The body holds the text of the user and, on its own line, the
 * planned attendees as text: they are never attendees of the event, so Graph
 * sends them nothing. Exchange adds a Teams join block at the end of the body
 * of an online meeting; the text of the user is the part before it. The body
 * of another event has no block, so all of it is text of the user.
 */

export const DRAFT_PREFIX = '[DRAFT] ';

export const PLANNED_MARK = 'Planned attendees (NOT invited):';

// The title of the Teams join block that Exchange adds at the end of the body of an online meeting.
const TEAMS_TITLE = 'Microsoft Teams meeting';

// The line of underscores that Exchange puts before the title of the block, at least this long.
const TEAMS_RULE = '_'.repeat(10);

/** The subject with the prefix `[DRAFT] ` once: a prefix that the user gave is not added again. The schema trims the subject. */
export const draftSubject = (subject: string): string => `${DRAFT_PREFIX}${subject.replace(/^\[DRAFT\]\s*/i, '')}`;

/** Whether the subject starts with the prefix that this CLI gives. */
export const hasDraftPrefix = (subject: string): boolean => subject.startsWith(DRAFT_PREFIX);

/** The planned attendees of an option: separated by `;` or `,`, each trimmed, empty ones dropped. */
export const plannedOf = (value: string): ReadonlyArray<string> =>
  value
    .split(/[;,]/)
    .map((item) => item.trim())
    .filter((item) => item !== '');

/** The body text: the text of the user, then the line of the planned attendees, if any. */
export const draftBodyText = (text: string, planned: ReadonlyArray<string>): string =>
  [text.trim(), planned.length > 0 ? `${PLANNED_MARK} ${planned.join('; ')}` : ''].filter((part) => part !== '').join('\n\n');

// The text before the index `at`, or all of it when `at` is -1 (not found).
const before = (text: string, at: number): string => (at === -1 ? text : text.slice(0, at));

// Where a line of underscores ends the text, or -1: the start of the block when it is just before the title.
const ruleStart = (head: string): number => {
  const text = head.trimEnd();
  const other = [...text].reverse().findIndex((char) => char !== '_');
  const rule = other === -1 ? text.length : other;
  return rule >= TEAMS_RULE.length ? text.length - rule : -1;
};

// Where the Teams block of a text body starts: the last title (the block is at the end), with its line of underscores; -1 if none.
const teamsStartOf = (bodyText: string): number => {
  const at = bodyText.lastIndexOf(TEAMS_TITLE);
  if (at === -1) return -1;
  const rule = ruleStart(bodyText.slice(0, at));
  return rule === -1 ? at : rule;
};

/**
 * The text of a body that Graph gave as text: the text of the user and the
 * planned attendees. Only an online meeting has a Teams join block to remove,
 * so the text of another event stays whole, whatever words it holds.
 */
export const draftTextOf = (bodyText: string, online: boolean): string => (online ? before(bodyText, teamsStartOf(bodyText)) : bodyText).trim();

/** The text of the user in a body that Graph gave as text, without the planned attendees and the Teams join block. */
export const ownTextOf = (bodyText: string, online: boolean): string => {
  const text = draftTextOf(bodyText, online);
  return before(text, text.indexOf(PLANNED_MARK));
};

/** The planned attendees on the line of the mark in a body that Graph gave as text. */
export const plannedInBody = (bodyText: string, online: boolean): ReadonlyArray<string> => {
  const text = draftTextOf(bodyText, online);
  const at = text.indexOf(PLANNED_MARK);
  if (at === -1) return [];
  const line = text.slice(at + PLANNED_MARK.length);
  return plannedOf(before(line, line.search(/\r?\n/)));
};

/** Whether two texts are the same when each run of white space is one space: Graph can change the line ends of a body. The texts are trimmed. */
export const sameText = (a: string, b: string): boolean => a.replaceAll(/\s+/g, ' ') === b.replaceAll(/\s+/g, ' ');

const escapeHtml = (text: string): string => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');

/**
 * The Teams join block of an HTML body, to its end, or an empty text: from
 * the `<div>` that holds the line of underscores before the last title, or
 * from the title when there is no such line.
 */
export const teamsBlockOf = (html: string): string => {
  const at = html.lastIndexOf(TEAMS_TITLE);
  if (at === -1) return '';
  const rule = html.lastIndexOf(TEAMS_RULE, at);
  const start = rule === -1 ? -1 : html.lastIndexOf('<div', rule);
  return html.slice(start === -1 ? at : start);
};

/** The HTML body of an online meeting: the text, then the Teams join block of before (to the end of the body), so Exchange keeps the meeting. */
export const htmlBodyWith = (text: string, block: string): string => `<html><body><p>${escapeHtml(text).replaceAll('\n', '<br>')}</p>${block}`;
