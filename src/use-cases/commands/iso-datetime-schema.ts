import { z } from 'zod';
import { parseIsoDateTime } from '../../domain/iso-datetime.ts';

/**
 * Reusable Zod field for an ISO-8601 UTC datetime parameter that ALSO accepts
 * the relative-date vocabulary defined in `src/domain/iso-datetime.ts`
 * (`7d`, `1w`, `today`, `monday`, `start-of-month`, etc.). Calendar commands
 * compose this into their schemas in place of `z.string().min(1)`; the
 * URL-builder sees the already-resolved canonical ISO form.
 *
 * removes the manual ISO arithmetic an LLM had to do
 * to ask "what changed this week" against `list-calendar-view`. Type
 * inference is left to Zod so the calendar commands' shape signatures stay
 * compatible with `buildListCommand`'s `ZodRawShape` constraint.
 */
export const isoDateTimeField = z
  .string()
  .min(1)
  .transform((value, ctx) => {
    const parsed = parseIsoDateTime(value);
    if (parsed.ok) return parsed.value;
    ctx.addIssue({
      code: 'custom',
      message: `not a recognised date format ("${parsed.error.input}"). ${parsed.error.hint}`,
    });
    return z.NEVER;
  });

/**
 * Standard `--start-date-time` / `--end-date-time` description used by every
 * calendar-view command. Mentions both the strict ISO form Graph expects and
 * the relative vocabulary the CLI also accepts, so the LLM sees both options
 * in `--help`.
 */
export const RELATIVE_DATE_DESCRIPTION =
  'ISO 8601 UTC (e.g. `2026-04-01T00:00:00Z` or `2026-04-01`) OR a relative shape: `7d` / `1w` / `2h` / `30m` (past), `+7d` (future), `today` / `yesterday` / `tomorrow` / `now`, `monday`-`sunday` (most recent), `last-<weekday>` / `next-<weekday>`, `start-of-week|month|year`, `end-of-week|month|year`. Relative forms resolve at request time relative to the CLI process clock (UTC).';
