import type { Result } from './result.ts';
import { err, ok } from './result.ts';

/**
 * Branded ISO-8601 UTC datetime string ("2026-04-01T00:00:00Z"). Constructed
 * only via `parseIsoDateTime` (or the unsafe escape hatch); every Graph URL
 * built from a calendar-window parameter accepts this type so an unvalidated
 * `string` can't slip through.
 *
 * the previous calendar commands accepted any
 * `z.string().min(1)`, so the LLM had to compute "last week" → ISO by hand.
 * `parseIsoDateTime` turns "7d" / "monday" / "today" / "2026-04-01" / a strict
 * ISO timestamp into the canonical ISO form, so the URL builders stay
 * unchanged.
 */
export type IsoDateTime = string & { readonly __brand: 'IsoDateTime' };

export const isoDateTimeUnsafe = (raw: string): IsoDateTime => raw as IsoDateTime;

export type DateParseError = {
  readonly type: 'invalid_format' | 'invalid_zone';
  readonly input: string;
  readonly hint: string;
};

const HINT = [
  'Accepted shapes:',
  '  - strict ISO 8601 UTC, e.g. `2026-04-01T00:00:00Z`',
  '  - ISO date, e.g. `2026-04-01` (expands to midnight UTC)',
  '  - past offset, e.g. `7d`, `1w`, `2h`, `30m`',
  '  - future offset, e.g. `+7d`, `+1w`',
  '  - named: `now`, `today`, `yesterday`, `tomorrow`',
  '  - weekday: `monday`-`sunday`, `last-monday`-`last-sunday`, `next-monday`-`next-sunday`',
  '  - boundary: `start-of-week|month|year`, `end-of-week|month|year` (week starts Monday)',
  "Named days and boundaries resolve in the run's time zone (the machine's, or `--tz` / `ASKMARCEL_TZ`); instants, bare dates and offsets do not depend on it.",
].join('\n');

const invalid = (input: string): DateParseError => ({ type: 'invalid_format', input, hint: HINT });
const ZONE_HINT = 'an IANA time zone such as `Europe/Amsterdam` or `Asia/Shanghai`, or `UTC`';
const invalidZone = (zone: string): DateParseError => ({ type: 'invalid_zone', input: zone, hint: ZONE_HINT });

const toIso = (d: Date): IsoDateTime => d.toISOString() as IsoDateTime;

// Strict ISO 8601 in UTC ("Z" suffix) — exactly the form Graph expects.
const STRICT_ISO_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,9})?Z$/;
const DATE_ONLY_RE = /^\d{4}-\d{2}-\d{2}$/;
const OFFSET_RE = /^([+-]?)(\d+)([dwhm])$/;

const WEEKDAY_INDEX: Readonly<Record<string, number>> = {
  monday: 1,
  tuesday: 2,
  wednesday: 3,
  thursday: 4,
  friday: 5,
  saturday: 6,
  sunday: 7,
};

const isoDateAtMidnightUtc = (raw: string): Result<IsoDateTime, DateParseError> => {
  // `Date.parse` refuses an impossible month and silently normalises an
  // impossible day (2026-02-30 becomes March 2); round-tripping the instant
  // through toISOString catches both, and a bare date stays midnight UTC.
  const t = Date.parse(`${raw}T00:00:00Z`);
  if (Number.isNaN(t)) return err(invalid(raw));
  const d = new Date(t);
  return d.toISOString().startsWith(raw) ? ok(toIso(d)) : err(invalid(raw));
};

type OffsetUnit = 'd' | 'w' | 'h' | 'm';

const UNIT_MS: Readonly<Record<OffsetUnit, number>> = {
  d: 86_400_000,
  w: 604_800_000,
  h: 3_600_000,
  m: 60_000,
};

const applyOffset = (now: Date, sign: 1 | -1, count: number, unit: OffsetUnit): IsoDateTime => {
  const ms = UNIT_MS[unit] * count * sign;
  return toIso(new Date(now.getTime() + ms));
};

// --- calendar arithmetic in a time zone -----------------------------------
//
// A named day (`today`, `monday`, `start-of-month`) is a wall-clock notion: it
// starts at midnight where the user is, not at midnight UTC. `Intl` gives the
// wall-clock parts of an instant in any IANA zone (the `clock` below reads
// them); the reverse (the instant at which a zone's clock reads 00:00 on a
// civil date) is the civil date taken as UTC minus the zone's offset at that
// moment, corrected once for a daylight change that may sit in between. With
// the zone `UTC` every formula collapses to the plain UTC arithmetic it
// replaces.

const CLOCK_FIELDS: Intl.DateTimeFormatOptions = { hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' };

/**
 * The wall-clock reader for `zone`. `Intl` resolves an IANA name, its older
 * aliases (`Asia/Kolkata` and `Asia/Calcutta`) and any letter case, which its
 * own list of zones does not, and throws a RangeError on anything else: a
 * native synchronous thrower, caught here and returned as a Result.
 */
const zoneClock = (zone: string): Result<Intl.DateTimeFormat, DateParseError> => {
  try {
    return ok(new Intl.DateTimeFormat('en-US', { ...CLOCK_FIELDS, timeZone: zone }));
  } catch {
    return err(invalidZone(zone));
  }
};

/** Whether `zone` names a time zone this runtime resolves: an IANA name or alias, `UTC`, in any letter case. */
export const isValidTimeZone = (zone: string): boolean => zoneClock(zone).ok;

type CivilDate = { readonly y: number; readonly m: number; readonly d: number };
type WallClock = CivilDate & { readonly h: number; readonly mi: number; readonly s: number };

const wallClock = (instant: Date, clock: Intl.DateTimeFormat): WallClock => {
  const parts: Record<string, number> = {};
  for (const part of clock.formatToParts(instant)) if (part.type !== 'literal') parts[part.type] = Number(part.value);
  return { y: parts['year'] ?? 0, m: parts['month'] ?? 0, d: parts['day'] ?? 0, h: parts['hour'] ?? 0, mi: parts['minute'] ?? 0, s: parts['second'] ?? 0 };
};

const zoneOffsetMs = (instant: Date, clock: Intl.DateTimeFormat): number => {
  const w = wallClock(instant, clock);
  return Date.UTC(w.y, w.m - 1, w.d, w.h, w.mi, w.s) - Math.floor(instant.getTime() / 1000) * 1000;
};

/** The instant at which the zone's clock reads 00:00 on the civil date; `d` may overflow the month, `Date.UTC` normalises it. */
const zoneMidnight = (civil: CivilDate, clock: Intl.DateTimeFormat): Date => {
  const asUtc = Date.UTC(civil.y, civil.m - 1, civil.d);
  const firstGuess = asUtc - zoneOffsetMs(new Date(asUtc), clock);
  return new Date(asUtc - zoneOffsetMs(new Date(firstGuess), clock));
};

const civilToday = (now: Date, clock: Intl.DateTimeFormat): CivilDate => {
  const w = wallClock(now, clock);
  return { y: w.y, m: w.m, d: w.d };
};

const shiftDays = (civil: CivilDate, days: number): CivilDate => ({ ...civil, d: civil.d + days });

// Mon=1 … Sun=7 (matches ISO 8601 weekday ordering). The weekday of a civil
// date does not depend on the zone, so the UTC calendar answers it.
const isoWeekday = (civil: CivilDate): number => {
  const js = new Date(Date.UTC(civil.y, civil.m - 1, civil.d)).getUTCDay();
  return js === 0 ? 7 : js;
};

type WeekdayDirection = 'this' | 'last' | 'next';

const directionDiff = (back: number, direction: WeekdayDirection): number => {
  if (direction === 'this') return -back;
  if (direction === 'last') return -back - 7;
  // next: when today matches the target, advance a full week; otherwise this
  // week's later occurrence.
  return back === 0 ? 7 : 7 - back;
};

const findWeekday = (now: Date, target: number, direction: WeekdayDirection, clock: Intl.DateTimeFormat): IsoDateTime => {
  const today = civilToday(now, clock);
  // `back` = how many days ago the target weekday last occurred, counting
  // today as 0. All three directions derive from it.
  //   this : today minus `back`            — most recent occurrence INCLUDING today
  //   last : today minus `back` minus 7    — the occurrence ONE FULL WEEK earlier
  //   next : today plus (7 - back)         — the occurrence after today (a full week when today matches)
  const back = (isoWeekday(today) - target + 7) % 7;
  return toIso(zoneMidnight(shiftDays(today, directionDiff(back, direction)), clock));
};

const startOfWeek = (now: Date, clock: Intl.DateTimeFormat): IsoDateTime => findWeekday(now, 1, 'this', clock);

const endOfWeek = (now: Date, clock: Intl.DateTimeFormat): IsoDateTime => {
  // End of week = the last millisecond before next Monday's midnight in the zone.
  const today = civilToday(now, clock);
  const back = (isoWeekday(today) - 1 + 7) % 7;
  return toIso(new Date(zoneMidnight(shiftDays(today, 7 - back), clock).getTime() - 1));
};

const startOfMonth = (now: Date, clock: Intl.DateTimeFormat): IsoDateTime => toIso(zoneMidnight({ ...civilToday(now, clock), d: 1 }, clock));
const endOfMonth = (now: Date, clock: Intl.DateTimeFormat): IsoDateTime => {
  const t = civilToday(now, clock);
  return toIso(new Date(zoneMidnight({ y: t.y, m: t.m + 1, d: 1 }, clock).getTime() - 1));
};
const startOfYear = (now: Date, clock: Intl.DateTimeFormat): IsoDateTime => toIso(zoneMidnight({ y: civilToday(now, clock).y, m: 1, d: 1 }, clock));
const endOfYear = (now: Date, clock: Intl.DateTimeFormat): IsoDateTime => toIso(new Date(zoneMidnight({ y: civilToday(now, clock).y + 1, m: 1, d: 1 }, clock).getTime() - 1));

const namedSimple = (now: Date, name: string, clock: Intl.DateTimeFormat): IsoDateTime | undefined => {
  if (name === 'now') return toIso(now);
  if (name === 'today') return toIso(zoneMidnight(civilToday(now, clock), clock));
  if (name === 'yesterday') return toIso(zoneMidnight(shiftDays(civilToday(now, clock), -1), clock));
  if (name === 'tomorrow') return toIso(zoneMidnight(shiftDays(civilToday(now, clock), 1), clock));
  if (name === 'start-of-week') return startOfWeek(now, clock);
  if (name === 'end-of-week') return endOfWeek(now, clock);
  if (name === 'start-of-month') return startOfMonth(now, clock);
  if (name === 'end-of-month') return endOfMonth(now, clock);
  if (name === 'start-of-year') return startOfYear(now, clock);
  if (name === 'end-of-year') return endOfYear(now, clock);
  return undefined;
};

/**
 * Parse a free-form date input into the canonical `2026-04-01T00:00:00Z`
 * shape (`IsoDateTime`). Returns `err(invalid_format)` with a multi-line
 * hint listing every accepted shape — surfaces directly through the CLI's
 * validation envelope so an LLM gets all the alternatives without an extra
 * round-trip. See `HINT` above for the full list. Named days and boundaries
 * resolve in `zone` (an IANA name, `UTC` by default); instants, bare dates and
 * offsets do not depend on it.
 */
export const parseIsoDateTime = (rawInput: string, now: Date = new Date(), zone: string = 'UTC'): Result<IsoDateTime, DateParseError> => {
  const clock = zoneClock(zone);
  if (!clock.ok) return clock;
  const input = rawInput.trim();
  if (input.length === 0) return err(invalid(rawInput));

  if (STRICT_ISO_RE.test(input)) return ok(input as IsoDateTime);
  if (DATE_ONLY_RE.test(input)) return isoDateAtMidnightUtc(input);

  const lower = input.toLowerCase();

  const simple = namedSimple(now, lower, clock.value);
  if (simple !== undefined) return ok(simple);

  const weekdayIdx = WEEKDAY_INDEX[lower];
  if (weekdayIdx !== undefined) return ok(findWeekday(now, weekdayIdx, 'this', clock.value));

  if (lower.startsWith('last-')) {
    const w = WEEKDAY_INDEX[lower.slice('last-'.length)];
    if (w !== undefined) return ok(findWeekday(now, w, 'last', clock.value));
  }
  if (lower.startsWith('next-')) {
    const w = WEEKDAY_INDEX[lower.slice('next-'.length)];
    if (w !== undefined) return ok(findWeekday(now, w, 'next', clock.value));
  }

  const offsetMatch = OFFSET_RE.exec(lower);
  if (offsetMatch !== null) {
    const [, signStr, countStr, unitStr] = offsetMatch;
    if (countStr === undefined || unitStr === undefined) return err(invalid(rawInput));
    const count = Number.parseInt(countStr, 10);
    if (Number.isNaN(count) || count < 0) return err(invalid(rawInput));
    // Default (no sign): treat as PAST offset, matching the "--since 7d" intent.
    const sign: 1 | -1 = signStr === '+' ? 1 : -1;
    const unit = unitStr as 'd' | 'w' | 'h' | 'm';
    return ok(applyOffset(now, sign, count, unit));
  }

  return err(invalid(rawInput));
};
