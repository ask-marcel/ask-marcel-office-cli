import { isDay, isMinute } from './todo-fields.ts';

/*
 * The start and the end of a calendar draft. A time is a local day and time
 * `YYYY-MM-DDTHH:mm` in the time zone of the mailbox, or a day `YYYY-MM-DD`
 * for an all-day event, whose end is the day after its last day.
 */

/** A start or an end in the Graph form. */
export type EventTime = { readonly dateTime: string; readonly timeZone: string };

/** Whether `value` has the form of a time of the event: a day for an all-day event, else a day and a time. */
export const isEventTime = (value: string, allDay: boolean): boolean => (allDay ? isDay(value) : isMinute(value));

/** The Graph form of a time that `isEventTime` accepted. */
export const eventTimeOf = (value: string, allDay: boolean, timeZone: string): EventTime => ({ dateTime: allDay ? `${value}T00:00:00` : `${value}:00`, timeZone });

/** The day and the minute of a Graph time (`2026-11-20T14:00:00.0000000` gives `2026-11-20T14:00`), to compare two times in one zone. */
export const minuteOf = (dateTime: string): string => dateTime.slice(0, 16);
