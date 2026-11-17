import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { err, ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import { createAttachmentUpload } from './attachment-upload.ts';
import type { FetchFn } from './graph-request.ts';
import { createGraphRequestCore } from './graph-request.ts';

/*
 * A file attached to a To Do task (Graph taskfileattachment-createuploadsession):
 * up to 25 MB. Under 3 MB in one POST; from 3 MB in an upload session on Graph
 * itself, so each byte range goes to `<uploadUrl>/content` WITH the bearer, and
 * the bearer goes to graph.microsoft.com over https only.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

const UPLOAD_URL = 'https://graph.microsoft.com/v1.0/users/u1/todo/lists/l1/tasks/t1/attachmentSessions/s1';
const TASK = { kind: 'todo-task', listId: 'l1', taskId: 't1' } as const;

type Sent = { readonly method: string; readonly url: string; readonly authorization: string | null; readonly range: string | null; readonly body: unknown };

const graph = (sent: Sent[], answer: (method: string, url: string) => Response): FetchFn => {
  const fetchFn: FetchFn = async (url, init) => {
    const headers = new Headers(init?.headers);
    const method = init?.method ?? 'GET';
    const body = typeof init?.body === 'string' ? JSON.parse(init.body) : undefined;
    sent.push({ method, url, authorization: headers.get('authorization'), range: headers.get('content-range'), body });
    return answer(method, url);
  };
  return fetchFn;
};

const uploadWith = (fetchFn: FetchFn): ReturnType<typeof createAttachmentUpload> => createAttachmentUpload(createGraphRequestCore(tokens, fetchFn), fetchFn);

describe('a small file attached to a To Do task', () => {
  it('is sent in one POST of a taskFileAttachment', async () => {
    const sent: Sent[] = [];
    const answer = await uploadWith(graph(sent, () => Response.json({ id: 'task-att-1' }, { status: 201 })))(TASK, { name: 'plan.md', bytes: new TextEncoder().encode('hi') });
    expect(answer).toEqual(ok({ id: 'task-att-1', name: 'plan.md', size: 2 }));
    expect(sent.map((s) => [s.method, s.url, s.body])).toEqual([
      [
        'POST',
        'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks/t1/attachments',
        { '@odata.type': '#microsoft.graph.taskFileAttachment', name: 'plan.md', contentBytes: 'aGk=' },
      ],
    ]);
  });
});

describe('a file of 3 MB or more attached to a To Do task', () => {
  it('opens an upload session and sends each signed range to the /content of the upload URL, then answers the attachment the Location names', async () => {
    const sent: Sent[] = [];
    const fetchFn = graph(sent, (method) => {
      if (method === 'POST') return Response.json({ uploadUrl: UPLOAD_URL });
      const puts = sent.filter((s) => s.method === 'PUT').length;
      return puts < 2
        ? Response.json({ nextExpectedRanges: ['3276800'] })
        : new Response(null, { status: 201, headers: { location: `${UPLOAD_URL.replace(/attachmentSessions\/s1$/, '')}Attachments/AAMk-task-att` } });
    });
    const answer = await uploadWith(fetchFn)(TASK, { name: 'scan.pdf', bytes: new Uint8Array(5_000_000), contentType: 'application/pdf' });
    expect(answer).toEqual(ok({ id: 'AAMk-task-att', name: 'scan.pdf', size: 5_000_000 }));
    expect(sent.map((s) => [s.method, s.url, s.authorization, s.range, s.body])).toEqual([
      [
        'POST',
        'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks/t1/attachments/createUploadSession',
        'Bearer basic-token',
        null,
        { attachmentInfo: { attachmentType: 'file', name: 'scan.pdf', size: 5_000_000, contentType: 'application/pdf' } },
      ],
      ['PUT', `${UPLOAD_URL}/content`, 'Bearer basic-token', 'bytes 0-3276799/5000000', undefined],
      ['PUT', `${UPLOAD_URL}/content`, 'Bearer basic-token', 'bytes 3276800-4999999/5000000', undefined],
    ]);
  });

  it('answers the attachment with no id when the last range comes back with no Location, since the file is attached', async () => {
    const fetchFn = graph([], (method) => (method === 'POST' ? Response.json({ uploadUrl: UPLOAD_URL }) : new Response(null, { status: 201 })));
    expect(await uploadWith(fetchFn)(TASK, { name: 'scan.pdf', bytes: new Uint8Array(3_000_000) })).toEqual(ok({ name: 'scan.pdf', size: 3_000_000 }));
  });

  it.each([UPLOAD_URL.replace('https:', 'http:'), 'https://graph.microsoft.com.example.net/s1', 'not a url'])(
    'does not send the bearer to an upload URL that is not Graph over https: %s',
    async (uploadUrl) => {
      const sent: Sent[] = [];
      const answer = await uploadWith(graph(sent, () => Response.json({ uploadUrl })))(TASK, { name: 'scan.pdf', bytes: new Uint8Array(3_000_000) });
      expect(answer.ok).toBe(false);
      if (answer.ok) return;
      expect(answer.error.type).toBe('network_error');
      expect(sent.map((s) => s.method)).toEqual(['POST']);
    }
  );

  it('cancels the upload session with a signed DELETE when Graph refuses a range', async () => {
    const sent: Sent[] = [];
    const fetchFn = graph(sent, (method) => {
      if (method === 'POST') return Response.json({ uploadUrl: UPLOAD_URL });
      return method === 'DELETE' ? new Response(null, { status: 204 }) : new Response(null, { status: 416 });
    });
    expect(await uploadWith(fetchFn)(TASK, { name: 'scan.pdf', bytes: new Uint8Array(3_000_000) })).toEqual(
      err({ type: 'api_error', status: 416, message: 'chunk PUT failed at byte 0' })
    );
    expect(sent.at(-1)).toEqual({ method: 'DELETE', url: UPLOAD_URL, authorization: 'Bearer basic-token', range: null, body: undefined });
  });

  it('still answers the refused range when the cancel of the session fails on the network', async () => {
    const sent: Sent[] = [];
    const fetchFn = graph(sent, (method) => {
      if (method === 'POST') return Response.json({ uploadUrl: UPLOAD_URL });
      if (method === 'DELETE') throw new Error('socket hang up');
      return new Response(null, { status: 416 });
    });
    expect(await uploadWith(fetchFn)(TASK, { name: 'scan.pdf', bytes: new Uint8Array(3_000_000) })).toEqual(
      err({ type: 'api_error', status: 416, message: 'chunk PUT failed at byte 0' })
    );
    expect(sent.map((s) => s.method)).toEqual(['POST', 'PUT', 'DELETE']);
  });

  it('fails when the upload session comes back with no upload URL', async () => {
    const answer = await uploadWith(graph([], () => Response.json({})))(TASK, { name: 'scan.pdf', bytes: new Uint8Array(3_000_000) });
    expect(answer).toEqual(err({ type: 'api_error', status: 500, message: 'createUploadSession returned no uploadUrl' }));
  });
});

describe('a file above the limit of To Do', () => {
  it('is refused before any request', async () => {
    const sent: Sent[] = [];
    const answer = await uploadWith(graph(sent, () => Response.json({})))(TASK, { name: 'video.mp4', bytes: new Uint8Array(25_000_001) });
    expect(answer).toEqual(
      err({ type: 'validation_error', code: 'attachment_too_large', message: 'video.mp4 is 25000001 bytes, above the limit of 25000000 bytes for an attachment of a To Do task.' })
    );
    expect(sent).toEqual([]);
  });

  it('is not refused at exactly 25 MB, the limit itself', async () => {
    const sent: Sent[] = [];
    await uploadWith(graph(sent, () => Response.json({})))(TASK, { name: 'video.mp4', bytes: new Uint8Array(25_000_000) });
    expect(sent.map((s) => s.url)).toEqual(['https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks/t1/attachments/createUploadSession']);
  });
});
