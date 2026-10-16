import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-request.ts';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { fakeCommand } from '../../test-helpers/command-package-fake.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import type { Command } from './command-types.ts';
import { withUnknownParamRejection } from './reject-unknown-params.ts';

/*
 * A parameter a command does not declare is refused before the command runs,
 * on every surface (core, package split phase 3): the Zod schemas strip unknown
 * keys, so a library or MCP caller would otherwise get data that ignored it.
 */

const folderCommand = (ran: string[]): Command =>
  withUnknownParamRejection({
    ...fakeCommand(),
    schema: z.object({ mailFolderId: z.string(), top: z.string().optional(), select: z.string().optional() }),
    execute: async (_graph, params) => {
      ran.push('graph');
      return ok(params);
    },
    executeLocal: async (_fs, params) => {
      ran.push('local');
      return ok(params);
    },
  });

const refusal = (message: string): Result<never, GraphError> => err({ type: 'validation_error', code: 'unknown_parameter', message });

describe('a parameter the command does not declare', () => {
  it('is refused before the request, with the nearest flag and the flags the command takes', async () => {
    const ran: string[] = [];
    expect(await folderCommand(ran).execute(fakeGraphClient(), { mailFolderId: 'inbox', selct: 'id' })).toEqual(
      refusal(
        '--selct is not a parameter of this command, so it would have been ignored rather than applied. Did you mean `--select`? Supported: --mail-folder-id, --select, --top.'
      )
    );
    expect(ran).toEqual([]);
  });

  it('names each unknown flag, in CLI spelling, with the suggestion for each', async () => {
    expect(await folderCommand([]).execute(fakeGraphClient(), { mailFolderId: 'inbox', selct: 'id', orderBy: 'x' })).toEqual(
      refusal(
        '--selct, --order-by are not parameters of this command, so it would have been ignored rather than applied. For `--selct`, did you mean `--select`? Supported: --mail-folder-id, --select, --top.'
      )
    );
  });

  it('suggests nothing when no flag is near', async () => {
    expect(await folderCommand([]).execute(fakeGraphClient(), { mailFolderId: 'inbox', zzzz: 'x' })).toEqual(
      refusal('--zzzz is not a parameter of this command, so it would have been ignored rather than applied. Supported: --mail-folder-id, --select, --top.')
    );
  });

  it('is refused on a local command too, which never reaches Graph', async () => {
    const ran: string[] = [];
    const result = await folderCommand(ran).executeLocal?.(createFileSystemFake(), { mailFolderId: 'inbox', bogus: 'x' });
    expect(result).toEqual(err(expect.objectContaining({ code: 'unknown_parameter' })));
    expect(ran).toEqual([]);
  });
});

describe('the parameters a command declares', () => {
  it('reach the command on Graph and on the local file system', async () => {
    const ran: string[] = [];
    const command = folderCommand(ran);
    expect(await command.execute(fakeGraphClient(), { mailFolderId: 'inbox', top: '5' })).toEqual(ok({ mailFolderId: 'inbox', top: '5' }));
    expect(await command.executeLocal?.(createFileSystemFake(), { mailFolderId: 'inbox' })).toEqual(ok({ mailFolderId: 'inbox' }));
    expect(ran).toEqual(['graph', 'local']);
  });

  it('are not checked on a command whose schema is not an object, which validates itself', async () => {
    const command = withUnknownParamRejection({ ...fakeCommand(), schema: z.string() });
    expect(await command.execute(fakeGraphClient(), { anything: 'x' })).toEqual(ok({ ranWith: { anything: 'x' } }));
  });

  it('leave a command with no local run without one', () => {
    expect(withUnknownParamRejection(fakeCommand()).executeLocal).toBeUndefined();
  });
});
