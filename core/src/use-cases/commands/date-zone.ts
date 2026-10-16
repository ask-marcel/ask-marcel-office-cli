/**
 * The time zone every date-taking command resolves named days in (`today`,
 * `monday`, `start-of-month`). The composition root sets it once per run from
 * `--tz`, `ASKMARCEL_TZ` or the machine; nothing here reads the environment,
 * so a test that never calls `setDateZone` runs in UTC, the domain default.
 */

let zone = 'UTC';

const setDateZone = (next: string): void => {
  zone = next;
};

const currentDateZone = (): string => zone;

export { currentDateZone, setDateZone };
