/*
 * The subject and the body of a calendar draft. The subject starts with
 * `[DRAFT] `. The body holds the text of the user and, on its own line, the
 * planned attendees as text: they are never attendees of the event, so Graph
 * sends them nothing. Exchange adds a Teams join block at the end of the body
 * of an online meeting; the text of the user is the part before it.
 */

export const DRAFT_PREFIX = '[DRAFT] ';

export const PLANNED_MARK = 'Planned attendees (NOT invited):';

// The start of the Teams join block in a text or an HTML body: a line of underscores, or its title.
const TEAMS_START = /_{10,}|Microsoft Teams meeting/;

/** The subject with the prefix `[DRAFT] ` once: a prefix that the user gave is not added again. */
export const draftSubject = (subject: string): string => `${DRAFT_PREFIX}${subject.trim().replace(/^\[DRAFT\]\s*/i, '')}`;

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

/** The text of a body that Graph gave as text, without the Teams join block: the text of the user and the planned attendees. */
export const draftTextOf = (bodyText: string): string => (bodyText.split(TEAMS_START)[0] ?? '').trim();

/** The text of the user in a body that Graph gave as text, without the planned attendees and the Teams join block. */
export const ownTextOf = (bodyText: string): string => (draftTextOf(bodyText).split(PLANNED_MARK)[0] ?? '').trim();

/** The planned attendees on the line of the mark in a body that Graph gave as text. */
export const plannedInBody = (bodyText: string): ReadonlyArray<string> => {
  const after = draftTextOf(bodyText).split(PLANNED_MARK)[1];
  return after === undefined ? [] : plannedOf(after.split(/\r?\n/)[0] ?? '');
};

/** The text with each run of white space as one space: Graph can change the line ends of a body. */
export const sameText = (a: string, b: string): boolean => a.replaceAll(/\s+/g, ' ').trim() === b.replaceAll(/\s+/g, ' ').trim();

const escapeHtml = (text: string): string => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');

/**
 * The Teams join block of an HTML body, to its end, or an empty text. A body
 * that Exchange converted from text has it from the title; an HTML body from
 * the `<div>` that holds the line of underscores.
 */
export const teamsBlockOf = (html: string): string => {
  const found = TEAMS_START.exec(html);
  if (found === null) return '';
  if (found[0].startsWith('Microsoft')) return html.slice(found.index);
  const start = html.lastIndexOf('<div', found.index);
  return start === -1 ? '' : html.slice(start);
};

/** The HTML body of an online meeting: the text, then the Teams join block of before (to the end of the body), so Exchange keeps the meeting. */
export const htmlBodyWith = (text: string, block: string): string => `<html><body><p>${escapeHtml(text).replaceAll('\n', '<br>')}</p>${block}`;
