import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { ReadCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import type { Converters } from '../ports/converters.ts';
import { withConverters } from './converting-command.ts';

const META = { summary: 's', category: 'files', graphMethod: 'GET', graphPathTemplate: '/x', options: [], effect: 'read' } as unknown as ReadCommand['meta'];
const SCHEMA = z.object({});
// A bag whose only member the commands below read is its marker.
const CONVERTERS = { stripHtmlToText: () => 'marker' } as unknown as Converters;

describe('a command and its converters', () => {
  it('builds the execute of a command on Graph from the converters', async () => {
    const command = withConverters({ meta: META, schema: SCHEMA, createExecute: (converters) => async () => ok(converters.stripHtmlToText('')) }, CONVERTERS);
    expect(Object.keys(command).toSorted((a, b) => a.localeCompare(b))).toEqual(['execute', 'meta', 'schema']);
    expect(await command.execute(fakeGraphClient(), {})).toEqual(ok('marker'));
  });

  it('builds the local execute of a command on a local file from the converters', async () => {
    const execute: ReadCommand['execute'] = async () => ok('graph');
    const command = withConverters({ meta: META, schema: SCHEMA, execute, createExecuteLocal: (converters) => async () => ok(converters.stripHtmlToText('')) }, CONVERTERS);
    expect(Object.keys(command).toSorted((a, b) => a.localeCompare(b))).toEqual(['execute', 'executeLocal', 'meta', 'schema']);
    expect(command.execute).toBe(execute);
    expect(await command.executeLocal?.({} as never, {})).toEqual(ok('marker'));
  });

  it('keeps a plain command as it is', () => {
    const plain: ReadCommand = { meta: META, schema: SCHEMA, execute: async () => ok('plain') };
    expect(withConverters(plain, CONVERTERS)).toBe(plain);
  });
});
