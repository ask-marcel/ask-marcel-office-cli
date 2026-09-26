import { isValidTimeZone } from '../domain/iso-datetime.ts';

/**
 * Which time zone a run resolves named days in: `--tz` when given (validated
 * by the option parser before it gets here), else `ASKMARCEL_TZ`, else the
 * machine's zone. A bad `ASKMARCEL_TZ` is not silently UTC: it falls back to
 * the machine with a warning the caller logs, and an unknown machine zone
 * (a stripped container) falls back to UTC.
 */

type ResolvedZone = { readonly zone: string; readonly source: 'flag' | 'env' | 'machine' | 'fallback'; readonly warning?: string };

const machineTimeZone = (): string => Intl.DateTimeFormat().resolvedOptions().timeZone;

const resolveDateZone = (explicit: string | undefined, env: string | undefined = process.env['ASKMARCEL_TZ'], machine: () => string = machineTimeZone): ResolvedZone => {
  if (explicit !== undefined && explicit !== '') return { zone: explicit, source: 'flag' };
  const fromMachine = machine();
  const machineZone: ResolvedZone = isValidTimeZone(fromMachine)
    ? { zone: fromMachine, source: 'machine' }
    : { zone: 'UTC', source: 'fallback', warning: `the machine reports an unknown time zone "${fromMachine}"; named days resolve in UTC` };
  if (env === undefined || env === '') return machineZone;
  if (isValidTimeZone(env)) return { zone: env, source: 'env' };
  return { ...machineZone, warning: `ASKMARCEL_TZ "${env}" is not a known IANA time zone; named days resolve in ${machineZone.zone}` };
};

export { machineTimeZone, resolveDateZone };
export type { ResolvedZone };
