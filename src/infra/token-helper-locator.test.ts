import { describe, expect, it } from 'bun:test';
import { err, ok } from '../domain/result.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '../test-helpers/filesystem-fake.ts';
import type { TokenHelperLocatorDeps } from './token-helper-locator.ts';
import { locateTokenHelper } from './token-helper-locator.ts';

const LOCATOR = '/home/me/.ask-marcel/token-helper.json';
const NODE = '/usr/local/bin/node';
const ENTRY = '/opt/ask-marcel/dist/token.js';
const WIN_NODE = 'C:\\Program Files\\nodejs\\node.exe';
const ON_PATH = { command: 'ask-marcel-office-auth', args: ['token'] };

// The shim npm writes for a global bin on Windows (cmd-shim), trimmed to the
// lines that matter: the program it runs is the quoted path before `%*`.
const NPM_SHIM = [
  '@ECHO off',
  'SETLOCAL',
  'CALL :find_dp0',
  'IF EXIST "%dp0%\\node.exe" (',
  '  SET "_prog=%dp0%\\node.exe"',
  ') ELSE (',
  '  SET "_prog=node"',
  ')',
  'endLocal & goto #_undefined_# 2>NUL || title %COMSPEC% & "%_prog%"  "%dp0%\\node_modules\\@ask-marcel\\office-auth\\dist\\cli.js" %*',
].join('\r\n');

const locate = (fs: FileSystemFake, extra: Partial<TokenHelperLocatorDeps> = {}): ReturnType<typeof locateTokenHelper> =>
  locateTokenHelper({ env: {}, locatorPath: LOCATOR, fs, platform: 'linux', execPath: NODE, ...extra });

const withFiles = (files: Record<string, string>): FileSystemFake => {
  const fs = createFileSystemFake();
  for (const [path, content] of Object.entries(files)) fs.seed(path, content);
  return fs;
};

describe('finding the token helper', () => {
  it('runs ASKMARCEL_TOKEN_COMMAND first, as the program itself with nothing before --tier, ahead of a locator file', async () => {
    const fs = withFiles({ [LOCATOR]: JSON.stringify({ execPath: NODE, entry: ENTRY, version: '2.8.0' }), [ENTRY]: '' });
    expect(await locate(fs, { env: { ASKMARCEL_TOKEN_COMMAND: '/opt/bin/token-helper' } })).toEqual(ok({ command: '/opt/bin/token-helper', args: [] }));
  });

  it('refuses an ASKMARCEL_TOKEN_COMMAND that is not an absolute path rather than look it up', async () => {
    expect(await locate(withFiles({}), { env: { ASKMARCEL_TOKEN_COMMAND: 'bin/token-helper' } })).toEqual(err({ type: 'not_absolute' }));
  });

  it('reads an absolute path on Windows by the Windows rules', async () => {
    const command = 'C:\\tools\\token-helper.exe';
    expect(await locate(withFiles({}), { platform: 'win32', env: { ASKMARCEL_TOKEN_COMMAND: command } })).toEqual(ok({ command, args: [] }));
  });

  it('runs the entry the locator file records with the runtime it records, so neither PATH nor a shebang matters', async () => {
    const fs = withFiles({ [LOCATOR]: JSON.stringify({ execPath: NODE, entry: ENTRY, version: '2.8.0' }), [ENTRY]: '#!/usr/bin/env node', [NODE]: '' });
    expect(await locate(fs)).toEqual(ok({ command: NODE, args: [ENTRY] }));
  });

  it('reads a locator file on Windows by the Windows rules, drive letters and all', async () => {
    const entry = 'C:\\Users\\me\\AppData\\Roaming\\npm\\node_modules\\@ask-marcel\\office-auth\\dist\\token.js';
    const fs = withFiles({ [LOCATOR]: JSON.stringify({ execPath: WIN_NODE, entry, version: '2.8.0' }), [entry]: '', [WIN_NODE]: 'MZ' });
    expect(await locate(fs, { platform: 'win32', env: {} })).toEqual(ok({ command: WIN_NODE, args: [entry] }));
  });

  it('passes over a locator file whose entry is no longer installed and looks on PATH', async () => {
    const fs = withFiles({ [LOCATOR]: JSON.stringify({ execPath: NODE, entry: ENTRY, version: '2.8.0' }), [NODE]: '' });
    expect(await locate(fs)).toEqual(ok(ON_PATH));
  });

  it('passes over a locator file whose runtime is gone (an upgraded node, a cleaned temporary runtime) and looks on PATH', async () => {
    const fs = withFiles({ [LOCATOR]: JSON.stringify({ execPath: NODE, entry: ENTRY, version: '2.8.0' }), [ENTRY]: '' });
    expect(await locate(fs)).toEqual(ok(ON_PATH));
  });

  it('passes over a locator file that does not hold a location', async () => {
    for (const held of [
      'null',
      '[]',
      '"text"',
      JSON.stringify({ execPath: NODE }),
      JSON.stringify({ execPath: 'node', entry: ENTRY }),
      JSON.stringify({ execPath: NODE, entry: 'dist/token.js' }),
      '{',
    ]) {
      const fs = withFiles({ [LOCATOR]: held, [ENTRY]: '', 'dist/token.js': '' });
      expect(await locate(fs)).toEqual(ok(ON_PATH));
    }
  });

  it('leaves the PATH search to the system off Windows: it names the auth bin and its token command', async () => {
    expect(await locate(withFiles({}))).toEqual(ok(ON_PATH));
  });

  it('runs the script behind an npm .cmd shim found on the Windows PATH with this runtime, never through a shell', async () => {
    const fs = withFiles({ 'C:\\npm\\ask-marcel-office-auth.cmd': NPM_SHIM });
    const found = await locate(fs, { platform: 'win32', execPath: WIN_NODE, env: { PATH: 'C:\\Windows;C:\\npm' } });
    expect(found).toEqual(ok({ command: WIN_NODE, args: ['C:\\npm\\node_modules\\@ask-marcel\\office-auth\\dist\\cli.js', 'token'] }));
  });

  it('reads the older shim form that names its folder as %~dp0', async () => {
    const fs = withFiles({ 'C:\\npm\\ask-marcel-office-auth.cmd': '@"%~dp0\\node.exe"  "%~dp0\\..\\lib\\node_modules\\auth\\cli.js" %*' });
    const found = await locate(fs, { platform: 'win32', execPath: WIN_NODE, env: { PATH: 'C:\\npm' } });
    expect(found).toEqual(ok({ command: WIN_NODE, args: ['C:\\lib\\node_modules\\auth\\cli.js', 'token'] }));
  });

  it('takes the first PATH folder that holds the bin, and in it the first extension PATHEXT lists', async () => {
    const fs = withFiles({ 'C:\\first\\ask-marcel-office-auth.exe': 'MZ', 'C:\\first\\ask-marcel-office-auth.cmd': NPM_SHIM, 'C:\\second\\ask-marcel-office-auth.exe': 'MZ' });
    const found = await locate(fs, { platform: 'win32', env: { PATH: 'C:\\first;C:\\second', PATHEXT: '.COM;.EXE;.BAT;.CMD' } });
    expect(found).toEqual(ok({ command: 'C:\\first\\ask-marcel-office-auth.exe', args: ['token'] }));
  });

  it('honours a PATHEXT that leaves .exe out', async () => {
    const fs = withFiles({ 'C:\\npm\\ask-marcel-office-auth.exe': 'MZ', 'C:\\npm\\ask-marcel-office-auth.cmd': NPM_SHIM });
    const found = await locate(fs, { platform: 'win32', execPath: WIN_NODE, env: { PATH: 'C:\\npm', PATHEXT: '.CMD' } });
    expect(found).toEqual(ok({ command: WIN_NODE, args: ['C:\\npm\\node_modules\\@ask-marcel\\office-auth\\dist\\cli.js', 'token'] }));
  });

  it('passes over a .bat or .cmd file that is not an npm shim, and finds nothing', async () => {
    const fs = withFiles({ 'C:\\npm\\ask-marcel-office-auth.bat': '@echo hello', 'C:\\npm\\ask-marcel-office-auth.cmd': 'start something %*' });
    expect(await locate(fs, { platform: 'win32', env: { PATH: 'C:\\npm' } })).toEqual(err({ type: 'not_found' }));
  });

  it('finds nothing on a Windows machine whose PATH holds no auth bin, or that has no PATH', async () => {
    expect(await locate(withFiles({}), { platform: 'win32', env: { PATH: 'C:\\Windows;;C:\\npm' } })).toEqual(err({ type: 'not_found' }));
    expect(await locate(withFiles({}), { platform: 'win32' })).toEqual(err({ type: 'not_found' }));
  });
});
