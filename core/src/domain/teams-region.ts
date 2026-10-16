import type { Result } from './result.ts';
import { err, ok } from './result.ts';

export type TeamsRegion = string & { readonly __brand: 'TeamsRegion' };

export type TeamsRegionError = { type: 'invalid_teams_region' };

/**
 * The regional segment of a Teams substrate URL
 * (`teams.microsoft.com/api/csa/<region>/...`, `.../api/chatsvc/<region>/...`).
 *
 * It is pasted into the path of every chat request, so it is checked once here
 * (hard rule 12): the shape the login capture accepts, and nothing that could
 * add a path segment or climb out of one. The host stays fixed either way; what
 * the check closes is a region steering a request to another path on it.
 */
const REGION_PATTERN = /^[a-z0-9-]+$/;

export const teamsRegion = (raw: string): Result<TeamsRegion, TeamsRegionError> => {
  if (!REGION_PATTERN.test(raw)) return err({ type: 'invalid_teams_region' });
  return ok(raw as TeamsRegion);
};
