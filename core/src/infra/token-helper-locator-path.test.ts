import { describe, expect, it } from 'bun:test';
import { resolveAuthPaths } from './auth-paths.ts';
import { tokenHelperLocatorPath } from './token-helper-locator.ts';

/*
 * Where the locator file of the token helper is: auth writes it and the
 * read and write packages read it, so both sides take the path from core.
 */
describe('the path of the token helper locator', () => {
  it('puts the locator next to the token cache, under the home folder', () => {
    expect(tokenHelperLocatorPath(resolveAuthPaths('/home/robin', {}))).toBe('/home/robin/.ask-marcel/token-helper.json');
  });

  it('keeps the locator next to the token cache when ASKMARCEL_BROWSER_PROFILE moves the browser profile', () => {
    expect(tokenHelperLocatorPath(resolveAuthPaths('/home/robin', { ASKMARCEL_BROWSER_PROFILE: '/data/profile' }))).toBe('/home/robin/.ask-marcel/token-helper.json');
  });
});
