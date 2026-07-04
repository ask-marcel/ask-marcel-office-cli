import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute, meta, schema } from './get-schedule.ts';

const fakeGraph = (overrides: Partial<GraphClient> = {}): GraphClient => fakeGraphClient(overrides);

type ScheduleBody = {
  readonly schedules: ReadonlyArray<string>;
  readonly startTime: { readonly dateTime: string; readonly timeZone: string };
  readonly endTime: { readonly dateTime: string; readonly timeZone: string };
  readonly availabilityViewInterval?: number;
};

describe('get-schedule', () => {
  it('returns validation_error when --schedules is missing', async () => {
    const result = await execute(fakeGraph(), { startDateTime: '2026-07-01T09:00:00Z', endDateTime: '2026-07-01T18:00:00Z' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('returns validation_error when a date bound is missing', async () => {
    const result = await execute(fakeGraph(), { schedules: 'alice@contoso.com', startDateTime: '2026-07-01T09:00:00Z' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('POSTs to /me/calendar/getSchedule with the emails, UTC-zoned bounds, and no Z suffix on the dateTime', async () => {
    let captured: ScheduleBody | undefined;
    const graph = fakeGraph({
      post: async (path, body) => {
        expect(path).toBe('/me/calendar/getSchedule');
        captured = body as ScheduleBody;
        return ok({ value: [] });
      },
    });

    await execute(graph, { schedules: 'alice@contoso.com', startDateTime: '2026-07-01T09:00:00Z', endDateTime: '2026-07-01T18:00:00Z' });

    expect(captured?.schedules).toEqual(['alice@contoso.com']);
    expect(captured?.startTime.timeZone).toBe('UTC');
    expect(captured?.endTime.timeZone).toBe('UTC');
    expect(captured?.startTime.dateTime).toMatch(/^2026-07-01T09:00:00/);
    expect(captured?.startTime.dateTime).not.toContain('Z');
  });

  it('splits a comma-separated --schedules list and trims surrounding whitespace', async () => {
    let captured: ScheduleBody | undefined;
    const graph = fakeGraph({
      post: async (_path, body) => {
        captured = body as ScheduleBody;
        return ok({ value: [] });
      },
    });

    await execute(graph, { schedules: 'alice@contoso.com , bob@contoso.com ,room-a@contoso.com', startDateTime: '2026-07-01', endDateTime: '2026-07-02' });

    expect(captured?.schedules).toEqual(['alice@contoso.com', 'bob@contoso.com', 'room-a@contoso.com']);
  });

  it('omits availabilityViewInterval from the body when the flag is absent', async () => {
    let captured: ScheduleBody | undefined;
    const graph = fakeGraph({
      post: async (_path, body) => {
        captured = body as ScheduleBody;
        return ok({ value: [] });
      },
    });

    await execute(graph, { schedules: 'alice@contoso.com', startDateTime: '2026-07-01', endDateTime: '2026-07-02' });

    expect(captured).not.toHaveProperty('availabilityViewInterval');
  });

  it('coerces --availability-view-interval to a number in the body when provided', async () => {
    let captured: ScheduleBody | undefined;
    const graph = fakeGraph({
      post: async (_path, body) => {
        captured = body as ScheduleBody;
        return ok({ value: [] });
      },
    });

    await execute(graph, { schedules: 'alice@contoso.com', startDateTime: '2026-07-01', endDateTime: '2026-07-02', availabilityViewInterval: '15' });

    expect(captured?.availabilityViewInterval).toBe(15);
  });

  it('returns validation_error when --schedules is only separators (no real address survives the split)', async () => {
    const result = await execute(fakeGraph(), { schedules: ' , ,', startDateTime: '2026-07-01', endDateTime: '2026-07-02' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('resolves the relative-date vocabulary before building the bounds', async () => {
    let captured: ScheduleBody | undefined;
    const graph = fakeGraph({
      post: async (_path, body) => {
        captured = body as ScheduleBody;
        return ok({ value: [] });
      },
    });

    await execute(graph, { schedules: 'alice@contoso.com', startDateTime: 'today', endDateTime: '+7d' });

    // `today` / `+7d` resolve to concrete ISO datetimes, not the literal tokens.
    expect(captured?.startTime.dateTime).toMatch(/^\d{4}-\d{2}-\d{2}T/);
    expect(captured?.startTime.dateTime).not.toBe('today');
    expect(captured?.endTime.dateTime).not.toBe('+7d');
  });

  it('propagates a Graph error from the POST unchanged', async () => {
    const apiError: GraphError = { type: 'api_error', status: 403, message: 'ErrorAccessDenied' };
    const graph = fakeGraph({ post: async () => err(apiError) });

    const result = await execute(graph, { schedules: 'alice@contoso.com', startDateTime: '2026-07-01', endDateTime: '2026-07-02' });

    expect(result).toEqual(err(apiError));
  });

  it('rejects a non-string schedules value at the schema level', () => {
    const parsed = schema.safeParse({ schedules: 42, startDateTime: '2026-07-01', endDateTime: '2026-07-02' });
    expect(parsed.success).toBe(false);
  });

  it('is a read-only POST (does not set the mutates flag) and documents Calendars.Read in its example', () => {
    expect(meta.mutates).toBeUndefined();
    expect(meta.graphMethod).toBe('POST');
    expect(meta.example).toContain('ask-marcel-office get-schedule');
  });
});
