import { join } from 'node:path';

/*
 * Where the CLI keeps its sign-in state, under one home folder. Composition once
 * read HOME first while auth read USERPROFILE first, so on a Windows shell that
 * sets both, the token cache and the browser profile could land under different
 * homes. Every caller now resolves both here, from the home it is handed
 * (`os.homedir()` in production, which is USERPROFILE on Windows and HOME
 * elsewhere).
 */
export type AuthPaths = { readonly tokenCache: string; readonly browserProfile: string };

export const resolveAuthPaths = (home: string, env: Readonly<Record<string, string | undefined>>): AuthPaths => {
  const base = join(home, '.ask-marcel');
  return {
    tokenCache: join(base, 'token-cache.json'),
    browserProfile: env['ASKMARCEL_BROWSER_PROFILE'] || join(base, 'browser-profile'),
  };
};
