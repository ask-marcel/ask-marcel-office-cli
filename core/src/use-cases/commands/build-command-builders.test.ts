import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import {
  buildCommand,
  buildElevatedCommand,
  buildElevatedListCommand,
  buildElevatedPickODataListCommand,
  buildElevatedSelectableCommand,
  buildFilterSelectListCommand,
  buildListCommand,
  buildNoSkipListCommand,
  buildPickODataListCommand,
  buildSelectableCommand,
} from './build-command.ts';
import { tenantIdShape } from './tenant-option.ts';

/*
 * The command builders (core, package split phase 3): each turns a path and a
 * schema into a command that checks its flags, adds the OData passthroughs it
 * advertises and the defaults it was given, and sends the GET on the right tier.
 */

const PARTNER = '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04';

type Call = { readonly tier: string; readonly path: string };

// A graph that records each GET and the tier it was sent on.
const recordingGraph = (calls: Call[]): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    get: async (path) => {
      calls.push({ tier: 'basic', path });
      return ok({ path });
    },
    getElevated: async (path) => {
      calls.push({ tier: 'elevated', path });
      return ok({ path });
    },
    getGuest: async (path, tenant) => {
      calls.push({ tier: `guest ${tenant}`, path });
      return ok({ path });
    },
  });

const itemSchema = z.object({ id: z.string().min(1), ...tenantIdShape });
const itemPath = (p: { readonly id: string }): string => `/items/${p.id}`;

const run = async (command: ReturnType<typeof buildCommand>, params: Record<string, string>): Promise<{ readonly calls: ReadonlyArray<Call>; readonly result: unknown }> => {
  const calls: Call[] = [];
  const result = await command.execute(recordingGraph(calls), params);
  return { calls, result };
};

const refused = { ok: false, error: { type: 'validation_error', message: '--id is missing' } };

describe('a builder on the home token', () => {
  it.each([
    ['buildCommand', buildCommand((p) => `/items/${p['id']}`, itemSchema)],
    ['buildListCommand', buildListCommand(itemPath, itemSchema)],
    ['buildSelectableCommand', buildSelectableCommand(itemPath, itemSchema)],
    ['buildFilterSelectListCommand', buildFilterSelectListCommand(itemPath, itemSchema)],
    ['buildNoSkipListCommand', buildNoSkipListCommand(itemPath, itemSchema)],
    ['buildPickODataListCommand', buildPickODataListCommand(itemPath, itemSchema, ['top'])],
  ])('%s sends a GET signed with the home token, or a guest token for a named partner tenant', async (_name, command) => {
    expect(await run(command, { id: '42' })).toEqual({ calls: [{ tier: 'basic', path: '/items/42' }], result: ok({ path: '/items/42' }) });
    expect((await run(command, { id: '42', tenantId: PARTNER })).calls).toEqual([{ tier: `guest ${PARTNER}`, path: '/items/42' }]);
  });

  it.each([
    ['buildCommand', buildCommand((p) => `/items/${p['id']}`, itemSchema)],
    ['buildElevatedCommand', buildElevatedCommand((p) => `/items/${p['id']}`, itemSchema)],
    ['buildListCommand', buildListCommand(itemPath, itemSchema)],
    ['buildElevatedListCommand', buildElevatedListCommand(itemPath, itemSchema)],
    ['buildSelectableCommand', buildSelectableCommand(itemPath, itemSchema)],
    ['buildElevatedSelectableCommand', buildElevatedSelectableCommand(itemPath, itemSchema)],
    ['buildFilterSelectListCommand', buildFilterSelectListCommand(itemPath, itemSchema)],
    ['buildNoSkipListCommand', buildNoSkipListCommand(itemPath, itemSchema)],
    ['buildPickODataListCommand', buildPickODataListCommand(itemPath, itemSchema, ['top'])],
    ['buildElevatedPickODataListCommand', buildElevatedPickODataListCommand(itemPath, itemSchema, ['top'])],
  ])('%s refuses a missing flag before any request', async (_name, command) => {
    expect(await run(command, {})).toEqual({ calls: [], result: refused });
  });
});

describe('a builder on the elevated token', () => {
  it.each([
    ['buildElevatedCommand', buildElevatedCommand((p) => `/chats/${p['id']}`, itemSchema)],
    ['buildElevatedListCommand', buildElevatedListCommand((p) => `/chats/${p.id}`, itemSchema)],
    ['buildElevatedSelectableCommand', buildElevatedSelectableCommand((p) => `/chats/${p.id}`, itemSchema)],
    ['buildElevatedPickODataListCommand', buildElevatedPickODataListCommand((p) => `/chats/${p.id}`, itemSchema, ['top'])],
  ])('%s signs with the elevated token, even when a partner tenant is named', async (_name, command) => {
    expect(await run(command, { id: '19:abc', tenantId: PARTNER })).toEqual({ calls: [{ tier: 'elevated', path: '/chats/19:abc' }], result: ok({ path: '/chats/19:abc' }) });
  });
});

describe('the OData passthroughs a builder advertises', () => {
  it('adds every OData flag of a list in a fixed order, encoded', async () => {
    const { calls } = await run(buildListCommand(itemPath, itemSchema), { id: '1', expand: 'a', orderby: 'b desc', filter: "c eq 'd'", select: 'e,f', skip: '2', top: '3' });
    expect(calls).toEqual([{ tier: 'basic', path: "/items/1?$top=3&$skip=2&$select=e%2Cf&$filter=c%20eq%20'd'&$orderby=b%20desc&$expand=a" }]);
  });

  it('adds the select and expand of a single resource, and refuses its other OData flags as unknown', async () => {
    const command = buildSelectableCommand(itemPath, itemSchema.strict());
    expect((await run(command, { id: '1', select: 'id', expand: 'x' })).calls).toEqual([{ tier: 'basic', path: '/items/1?$select=id&$expand=x' }]);
    expect((await run(command, { id: '1', top: '3' })).calls).toEqual([]);
  });

  it('adds the filter and select of a filter-select list, and refuses a top', async () => {
    const command = buildFilterSelectListCommand(itemPath, itemSchema.strict());
    expect((await run(command, { id: '1', filter: 'x', select: 'y' })).calls).toEqual([{ tier: 'basic', path: '/items/1?$select=y&$filter=x' }]);
    expect((await run(command, { id: '1', top: '3' })).calls).toEqual([]);
  });

  it('refuses a skip on a list that cannot skip, and keeps the other flags', async () => {
    const command = buildNoSkipListCommand(itemPath, itemSchema.strict());
    expect((await run(command, { id: '1', top: '3', orderby: 'x' })).calls).toEqual([{ tier: 'basic', path: '/items/1?$top=3&$orderby=x' }]);
    expect((await run(command, { id: '1', skip: '3' })).calls).toEqual([]);
  });

  it('takes only the flags a picked list names', async () => {
    const basic = buildPickODataListCommand(itemPath, itemSchema.strict(), ['top', 'select']);
    expect((await run(basic, { id: '1', top: '3', select: 'id' })).calls).toEqual([{ tier: 'basic', path: '/items/1?$top=3&$select=id' }]);
    expect((await run(basic, { id: '1', filter: 'x' })).calls).toEqual([]);
    const elevated = buildElevatedPickODataListCommand(itemPath, itemSchema.strict(), ['top']);
    expect((await run(elevated, { id: '1', top: '3' })).calls).toEqual([{ tier: 'elevated', path: '/items/1?$top=3' }]);
    expect((await run(elevated, { id: '1', select: 'x' })).calls).toEqual([]);
  });

  it('adds the select and expand of an elevated single resource', async () => {
    const command = buildElevatedSelectableCommand(itemPath, itemSchema);
    expect((await run(command, { id: '1', select: 'topic', expand: 'members' })).calls).toEqual([{ tier: 'elevated', path: '/items/1?$select=topic&$expand=members' }]);
  });

  it('adds every OData flag of an elevated list', async () => {
    const command = buildElevatedListCommand(itemPath, itemSchema);
    expect((await run(command, { id: '1', top: '5', filter: 'x' })).calls).toEqual([{ tier: 'elevated', path: '/items/1?$top=5&$filter=x' }]);
  });
});

describe('the defaults a builder is given', () => {
  const builders = [
    ['buildListCommand', (select: string) => buildListCommand(itemPath, itemSchema, { defaultSelect: select })],
    ['buildElevatedListCommand', (select: string) => buildElevatedListCommand(itemPath, itemSchema, { defaultSelect: select })],
    ['buildSelectableCommand', (select: string) => buildSelectableCommand(itemPath, itemSchema, { defaultSelect: select })],
    ['buildElevatedSelectableCommand', (select: string) => buildElevatedSelectableCommand(itemPath, itemSchema, { defaultSelect: select })],
    ['buildFilterSelectListCommand', (select: string) => buildFilterSelectListCommand(itemPath, itemSchema, { defaultSelect: select })],
    ['buildNoSkipListCommand', (select: string) => buildNoSkipListCommand(itemPath, itemSchema, { defaultSelect: select })],
    ['buildPickODataListCommand', (select: string) => buildPickODataListCommand(itemPath, itemSchema, ['select'], { defaultSelect: select })],
    ['buildElevatedPickODataListCommand', (select: string) => buildElevatedPickODataListCommand(itemPath, itemSchema, ['select'], { defaultSelect: select })],
  ] as const;

  it.each(builders)('%s selects the default fields when the caller selects none, and the caller’s fields otherwise', async (_name, build) => {
    const command = build('id,subject');
    expect((await run(command, { id: '1' })).calls[0]?.path).toBe('/items/1?$select=id%2Csubject');
    expect((await run(command, { id: '1', select: 'from' })).calls[0]?.path).toBe('/items/1?$select=from');
  });

  it.each([
    ['buildListCommand', buildListCommand(itemPath, itemSchema)],
    ['buildElevatedListCommand', buildElevatedListCommand(itemPath, itemSchema)],
    ['buildSelectableCommand', buildSelectableCommand(itemPath, itemSchema)],
    ['buildElevatedSelectableCommand', buildElevatedSelectableCommand(itemPath, itemSchema)],
    ['buildFilterSelectListCommand', buildFilterSelectListCommand(itemPath, itemSchema)],
    ['buildNoSkipListCommand', buildNoSkipListCommand(itemPath, itemSchema)],
    ['buildPickODataListCommand', buildPickODataListCommand(itemPath, itemSchema, ['select'])],
    ['buildElevatedPickODataListCommand', buildElevatedPickODataListCommand(itemPath, itemSchema, ['select'])],
  ])('%s adds no query when it has no default and the caller passes no flag', async (_name, command) => {
    expect((await run(command, { id: '1' })).calls[0]?.path).toBe('/items/1');
  });

  it('pages a list by the default top when the caller names none, and by the caller’s top otherwise', async () => {
    const command = buildListCommand(itemPath, itemSchema, { defaultTop: '25', defaultSelect: 'id' });
    expect((await run(command, { id: '1' })).calls[0]?.path).toBe('/items/1?$top=25&$select=id');
    expect((await run(command, { id: '1', top: '5' })).calls[0]?.path).toBe('/items/1?$top=5&$select=id');
  });

  it('keeps a query the path already carries, and joins the OData flags to it', async () => {
    const command = buildListCommand((p) => `/items/${p.id}/search(q='x')?scope=all`, itemSchema);
    expect((await run(command, { id: '1', top: '2' })).calls[0]?.path).toBe("/items/1/search(q='x')?scope=all&$top=2");
  });
});
