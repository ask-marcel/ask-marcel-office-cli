import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import { resolveAuthPaths } from './auth-paths.ts';
import { tokenHelperLocatorPath } from './token-helper-locator.ts';

/*
 * Where the locator file of the token helper is: auth writes it and the
 * read and write packages read it, so both sides take the path from core.
 */

// The separator is the one of the system that runs the CLI: a backslash on Windows.
const LOCATOR = join('/home/robin', '.ask-marcel', 'token-helper.json');

describe('the path of the token helper locator', () => {
  it('puts the locator next to the token cache, under the home folder', () => {
    expect(tokenHelperLocatorPath(resolveAuthPaths('/home/robin', {}))).toBe(LOCATOR);
  });

  it('keeps the locator next to the token cache when ASKMARCEL_BROWSER_PROFILE moves the browser profile', () => {
    expect(tokenHelperLocatorPath(resolveAuthPaths('/home/robin', { ASKMARCEL_BROWSER_PROFILE: '/data/profile' }))).toBe(LOCATOR);
  });
});
