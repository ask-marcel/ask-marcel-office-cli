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
  readonly type: 'invalid_format';
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
  '  - boundary: `start-of-week|month|year`, `end-of-week|month|year` (UTC, week starts Monday)',
].join('\n');

const invalid = (input: string): DateParseError => ({ type: 'invalid_format', input, hint: HINT });

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
  // `Date.parse('2026-04-01')` returns midnight UTC in current JS — but we go
  // through `new Date(year, monthIdx, day)` with explicit UTC builder to keep
  // it timezone-stable regardless of where the test runs.
  const parts = raw.split('-').map((p) => Number.parseInt(p, 10));
  const yearPart = parts[0];
  const monthPart = parts[1];
  const dayPart = parts[2];
  if (yearPart === undefined || monthPart === undefined || dayPart === undefined) return err(invalid(raw));
  if ([yearPart, monthPart, dayPart].some((n) => Number.isNaN(n))) return err(invalid(raw));
  const d = new Date(Date.UTC(yearPart, monthPart - 1, dayPart, 0, 0, 0, 0));
  // Guard against rollovers: `new Date(2026, 1, 30)` becomes March 2 silently.
  if (d.getUTCFullYear() !== yearPart || d.getUTCMonth() + 1 !== monthPart || d.getUTCDate() !== dayPart) return err(invalid(raw));
  return ok(toIso(d));
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

const startOfUtcDay = (d: Date): Date => new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), 0, 0, 0, 0));

// Mon=1 … Sun=7 (matches ISO 8601 weekday ordering). JS `getUTCDay()` returns
// 0 (Sun) … 6 (Sat); rotate so Mon is 1.
const isoWeekday = (d: Date): number => {
  const js = d.getUTCDay();
  return js === 0 ? 7 : js;
};

type WeekdayDirection = 'this' | 'last' | 'next';

const directionDiff = (back: number, direction: WeekdayDirection): number => {
  if (direction === 'this') return -back;
  if (direction === 'last') return -back - 7;
  // next: when today matches the target, advance a full week; otherwise this
  // week's later occurrence.
  if (back === 0) return 7;
  return 7 - back;
};

const findWeekday = (now: Date, target: number, direction: WeekdayDirection): IsoDateTime => {
  const today = startOfUtcDay(now);
  const todayIdx = isoWeekday(today);
  // `back` is the day count back to the most recent occurrence of `target`
  // (0 when today matches). All three directions derive from it.
  //   this : today minus `back`            — most recent occurrence INCLUDING today
  //   last : today minus `back` minus 7    — the occurrence ONE FULL WEEK earlier
  //   next : `back == 0` means 7 forward (next week's same weekday); else
  //          7 minus `back` (this week's later occurrence)
  const back = (todayIdx - target + 7) % 7;
  const diff = directionDiff(back, direction);
  return toIso(new Date(today.getTime() + diff * 86_400_000));
};

const startOfWeek = (now: Date): IsoDateTime => findWeekday(now, 1, 'this');

const endOfWeek = (now: Date): IsoDateTime => {
  // End of week = next Sunday at 23:59:59.999 UTC; computed from start-of-week
  // + 7 days - 1ms so it shares the same rounding rule as the others.
  const start = new Date(startOfWeek(now));
  return toIso(new Date(start.getTime() + 7 * 86_400_000 - 1));
};

const startOfMonth = (now: Date): IsoDateTime => toIso(new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1, 0, 0, 0, 0)));

const endOfMonth = (now: Date): IsoDateTime => toIso(new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1, 0, 0, 0, 0) - 1));

const startOfYear = (now: Date): IsoDateTime => toIso(new Date(Date.UTC(now.getUTCFullYear(), 0, 1, 0, 0, 0, 0)));

const endOfYear = (now: Date): IsoDateTime => toIso(new Date(Date.UTC(now.getUTCFullYear() + 1, 0, 1, 0, 0, 0, 0) - 1));

const namedSimple = (now: Date, name: string): IsoDateTime | undefined => {
  if (name === 'now') return toIso(now);
  if (name === 'today') return toIso(startOfUtcDay(now));
  if (name === 'yesterday') return toIso(new Date(startOfUtcDay(now).getTime() - 86_400_000));
  if (name === 'tomorrow') return toIso(new Date(startOfUtcDay(now).getTime() + 86_400_000));
  if (name === 'start-of-week') return startOfWeek(now);
  if (name === 'end-of-week') return endOfWeek(now);
  if (name === 'start-of-month') return startOfMonth(now);
  if (name === 'end-of-month') return endOfMonth(now);
  if (name === 'start-of-year') return startOfYear(now);
  if (name === 'end-of-year') return endOfYear(now);
  return undefined;
};

/**
 * Parse a free-form date input into the canonical `2026-04-01T00:00:00Z`
 * shape (`IsoDateTime`). Returns `err(invalid_format)` with a multi-line
 * hint listing every accepted shape — surfaces directly through the CLI's
 * validation envelope so an LLM gets all the alternatives without an extra
 * round-trip. See `HINT` above for the full list.
 */
export const parseIsoDateTime = (rawInput: string, now: Date = new Date()): Result<IsoDateTime, DateParseError> => {
  const input = rawInput.trim();
  if (input.length === 0) return err(invalid(rawInput));

  if (STRICT_ISO_RE.test(input)) return ok(input as IsoDateTime);
  if (DATE_ONLY_RE.test(input)) return isoDateAtMidnightUtc(input);

  const lower = input.toLowerCase();

  const simple = namedSimple(now, lower);
  if (simple !== undefined) return ok(simple);

  const weekdayIdx = WEEKDAY_INDEX[lower];
  if (weekdayIdx !== undefined) return ok(findWeekday(now, weekdayIdx, 'this'));

  if (lower.startsWith('last-')) {
    const w = WEEKDAY_INDEX[lower.slice('last-'.length)];
    if (w !== undefined) return ok(findWeekday(now, w, 'last'));
  }
  if (lower.startsWith('next-')) {
    const w = WEEKDAY_INDEX[lower.slice('next-'.length)];
    if (w !== undefined) return ok(findWeekday(now, w, 'next'));
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
