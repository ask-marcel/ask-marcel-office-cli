import { posix, win32 } from 'node:path';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';

/*
 * Where the token helper is, in the order the package split fixes
 * (docs/plans/2026-10-01-package-split.md, Token protocol):
 *
 * 1. ASKMARCEL_TOKEN_COMMAND, an absolute path, run as it is: never handed to a
 *    shell, so nothing in it is parsed.
 * 2. The locator file auth writes on every run (`{ execPath, entry, version }`):
 *    `execPath entry` runs whatever PATH holds and whether or not the entry
 *    keeps its shebang, which also covers npx and GUI-launched MCP clients. A
 *    file whose entry or runtime is gone (an evicted npx cache, an upgraded
 *    node) is passed over. The entry is checked first: it is the small file.
 * 3. The auth bin on PATH. Off Windows the system searches PATH when the child
 *    starts. On Windows a global npm bin is a `.cmd` shim, which only a shell
 *    runs, so the PATHEXT search is done here and a shim is read for the script
 *    it starts, which then runs with this process's own runtime.
 */

// How to run the token helper: a program and the arguments that come before
// `--tier`. Built only here, from the checked paths above.
export type TokenHelperCommand = { readonly command: string; readonly args: ReadonlyArray<string> };

export type TokenHelperLocateError = { readonly type: 'not_found' } | { readonly type: 'not_absolute' };

export type TokenHelperLocatorDeps = {
  readonly env: Readonly<Record<string, string | undefined>>;
  readonly locatorPath: string;
  readonly fs: Pick<FileSystem, 'readJson' | 'readBytes'>;
  readonly platform: string;
  // The runtime that runs the script behind a Windows shim: this process's own.
  readonly execPath: string;
};

const AUTH_BIN = 'ask-marcel-office-auth';
const TOKEN_ARGS = ['token'];
const DEFAULT_PATHEXT = '.COM;.EXE;.BAT;.CMD';
// The script a cmd-shim starts: the quoted `%dp0%\...` (or older `%~dp0\...`)
// path just before `%*`.
const SHIM_SCRIPT = /"%(?:~dp0|dp0%)\\([^"]+)"\s+%\*/;

type Located = Result<TokenHelperCommand, TokenHelperLocateError> | undefined;

const exists = async (deps: TokenHelperLocatorDeps, path: string): Promise<boolean> => (await deps.fs.readBytes(path)).ok;

const fromLocatorFile = async (deps: TokenHelperLocatorDeps, paths: typeof posix): Promise<Located> => {
  const held = await deps.fs.readJson<unknown>(deps.locatorPath);
  if (!held.ok || typeof held.value !== 'object' || held.value === null) return undefined;
  const { execPath, entry } = held.value as Partial<Record<'execPath' | 'entry', unknown>>;
  if (typeof execPath !== 'string' || typeof entry !== 'string' || !paths.isAbsolute(execPath) || !paths.isAbsolute(entry)) return undefined;
  return (await exists(deps, entry)) && (await exists(deps, execPath)) ? ok({ command: execPath, args: [entry] }) : undefined;
};

// A `.cmd` or `.bat` file runs only through a shell, so only an npm shim, read
// for its script, is taken; `.exe` and `.com` run as they are.
const runnable = async (deps: TokenHelperLocatorDeps, path: string): Promise<TokenHelperCommand | undefined> => {
  const bytes = await deps.fs.readBytes(path);
  if (!bytes.ok) return undefined;
  if (/\.(?:exe|com)$/i.test(path)) return { command: path, args: TOKEN_ARGS };
  const script = SHIM_SCRIPT.exec(new TextDecoder().decode(bytes.value))?.[1];
  return script === undefined ? undefined : { command: deps.execPath, args: [win32.join(win32.dirname(path), script), ...TOKEN_ARGS] };
};

const onWindowsPath = async (deps: TokenHelperLocatorDeps): Promise<Result<TokenHelperCommand, TokenHelperLocateError>> => {
  const folders = (deps.env['PATH'] ?? '').split(';').filter((folder) => folder !== '');
  const extensions = (deps.env['PATHEXT'] ?? DEFAULT_PATHEXT).split(';').filter((ext) => /^\.(?:exe|com|bat|cmd)$/i.test(ext));
  for (const candidate of folders.flatMap((folder) => extensions.map((ext) => win32.join(folder, `${AUTH_BIN}${ext.toLowerCase()}`)))) {
    const found = await runnable(deps, candidate);
    if (found !== undefined) return ok(found);
  }
  return err({ type: 'not_found' });
};

export const locateTokenHelper = async (deps: TokenHelperLocatorDeps): Promise<Result<TokenHelperCommand, TokenHelperLocateError>> => {
  const paths = deps.platform === 'win32' ? win32 : posix;
  const command = deps.env['ASKMARCEL_TOKEN_COMMAND'];
  if (command) return paths.isAbsolute(command) ? ok({ command, args: [] }) : err({ type: 'not_absolute' });
  const recorded = await fromLocatorFile(deps, paths);
  if (recorded !== undefined) return recorded;
  if (deps.platform !== 'win32') return ok({ command: AUTH_BIN, args: TOKEN_ARGS });
  return onWindowsPath(deps);
};
