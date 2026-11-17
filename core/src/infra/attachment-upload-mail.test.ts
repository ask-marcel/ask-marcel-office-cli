import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { err, ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import { createAttachmentUpload } from './attachment-upload.ts';
import type { FetchFn } from './graph-request.ts';
import { createGraphRequestCore } from './graph-request.ts';

/*
 * A file attached to a mail draft (Graph attachment-createuploadsession): under
 * 3 MB in one POST of the base64 bytes; from 3 MB in an upload session, whose
 * pre-authenticated Outlook URL takes the byte ranges with no Authorization
 * header (under 4 MB each), answers 200 while bytes remain and 201 with a
 * Location that names the attachment when the file is complete.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

const UPLOAD_URL = "https://outlook.office.com/api/v2.0/Users('u1')/Messages('m1')/AttachmentSessions('s1')?authtoken=opaque";
const LOCATION = "https://outlook.office.com/api/v2.0/Users('u1')/Messages('m1')/Attachments('AAMk-att=')";
const DRAFT = { kind: 'mail-draft', messageId: 'm1' } as const;

type Sent = {
  readonly method: string;
  readonly url: string;
  readonly authorization: string | null;
  readonly type: string | null;
  readonly range: string | null;
  readonly body: unknown;
};

// Graph answers each request with `answer`, which sees its method, URL and byte range.
const graph = (sent: Sent[], answer: (method: string, url: string, range: string | null) => Response): FetchFn => {
  const fetchFn: FetchFn = async (url, init) => {
    const headers = new Headers(init?.headers);
    const method = init?.method ?? 'GET';
    const body = typeof init?.body === 'string' ? JSON.parse(init.body) : init?.body;
    sent.push({ method, url, authorization: headers.get('authorization'), type: headers.get('content-type'), range: headers.get('content-range'), body });
    return answer(method, url, headers.get('content-range'));
  };
  return fetchFn;
};

const uploadWith = (fetchFn: FetchFn): ReturnType<typeof createAttachmentUpload> => createAttachmentUpload(createGraphRequestCore(tokens, fetchFn), fetchFn);

// An Outlook that opens a session at UPLOAD_URL, accepts each range with 200 and completes on the range that ends the file.
const outlook = (method: string, _url: string, range: string | null): Response => {
  if (method === 'POST') return Response.json({ uploadUrl: UPLOAD_URL, nextExpectedRanges: ['0-'] }, { status: 201 });
  const [, end, total] = /-(\d+)\/(\d+)$/.exec(range ?? '') ?? [];
  return Number(end) + 1 === Number(total)
    ? new Response(null, { status: 201, headers: { location: LOCATION } })
    : Response.json({ nextExpectedRanges: [String(Number(end) + 1)] });
};

describe('a small file attached to a mail draft', () => {
  it('is sent in one POST of its base64 bytes, and answers the attachment without the bytes Graph echoes', async () => {
    const sent: Sent[] = [];
    const fetchFn = graph(sent, () => Response.json({ id: 'att-1', name: 'notes.txt', size: 5, contentBytes: 'aGVsbG8=' }, { status: 201 }));
    const answer = await uploadWith(fetchFn)(DRAFT, { name: 'notes.txt', bytes: new TextEncoder().encode('hello'), contentType: 'text/plain' });
    expect(answer).toEqual(ok({ id: 'att-1', name: 'notes.txt', size: 5 }));
    expect(sent).toEqual([
      {
        method: 'POST',
        url: 'https://graph.microsoft.com/v1.0/me/messages/m1/attachments',
        authorization: 'Bearer basic-token',
        type: 'application/json',
        range: null,
        body: { '@odata.type': '#microsoft.graph.fileAttachment', name: 'notes.txt', contentBytes: 'aGVsbG8=', contentType: 'text/plain' },
      },
    ]);
  });
});

describe('a mail draft id that would turn the upload into a send', () => {
  it('is refused before any request', async () => {
    const sent: Sent[] = [];
    const answer = await uploadWith(graph(sent, () => Response.json({ id: 'att-1' })))(
      { kind: 'mail-draft', messageId: '../../me/sendMail?' },
      { name: 'a.txt', bytes: new Uint8Array(1) }
    );
    expect(sent).toEqual([]);
    expect(answer.ok).toBe(false);
    if (answer.ok) return;
    expect(answer.error.code).toBe('send_refused');
  });
});

describe('a file of 3 MB or more attached to a mail draft', () => {
  it('opens an upload session and sends unsigned ranges under 4 MB, then answers the attachment the Location names', async () => {
    const sent: Sent[] = [];
    const total = 7_000_000;
    const answer = await uploadWith(graph(sent, outlook))(DRAFT, { name: 'deck.pptx', bytes: new Uint8Array(total) });
    expect(answer).toEqual(ok({ id: 'AAMk-att=', name: 'deck.pptx', size: total }));
    expect(sent[0]).toEqual({
      method: 'POST',
      url: 'https://graph.microsoft.com/v1.0/me/messages/m1/attachments/createUploadSession',
      authorization: 'Bearer basic-token',
      type: 'application/json',
      range: null,
      body: { AttachmentItem: { attachmentType: 'file', name: 'deck.pptx', size: total } },
    });
    expect(sent.slice(1).map((s) => [s.method, s.url, s.authorization, s.type, s.range])).toEqual([
      ['PUT', UPLOAD_URL, null, 'application/octet-stream', 'bytes 0-3276799/7000000'],
      ['PUT', UPLOAD_URL, null, 'application/octet-stream', 'bytes 3276800-6553599/7000000'],
      ['PUT', UPLOAD_URL, null, 'application/octet-stream', 'bytes 6553600-6999999/7000000'],
    ]);
  });

  it('is sent in one POST when Outlook says the file is below the minimum size of an upload session', async () => {
    const sent: Sent[] = [];
    const fetchFn = graph(sent, (method, url) => {
      if (url.endsWith('/createUploadSession')) {
        return Response.json({ error: { code: 'ErrorAttachmentSizeShouldNotBeLessThanMinimumSize', message: 'too small' } }, { status: 400 });
      }
      return Response.json({ id: 'att-2' }, { status: 201 });
    });
    const answer = await uploadWith(fetchFn)(DRAFT, { name: 'scan.pdf', bytes: new Uint8Array(3_100_000) });
    expect(answer).toEqual(ok({ id: 'att-2', name: 'scan.pdf', size: 3_100_000 }));
    expect(sent.map((s) => `${s.method} ${s.url}`)).toEqual([
      'POST https://graph.microsoft.com/v1.0/me/messages/m1/attachments/createUploadSession',
      'POST https://graph.microsoft.com/v1.0/me/messages/m1/attachments',
    ]);
  });

  it.each([
    ['https://upload.example.com/session', 'upload.example.com'],
    [UPLOAD_URL.replace('https:', 'http:'), 'outlook.office.com'],
  ])('is not sent to an upload URL outside Outlook over https: %s', async (uploadUrl, host) => {
    const sent: Sent[] = [];
    const fetchFn = graph(sent, () => Response.json({ uploadUrl }, { status: 201 }));
    const answer = await uploadWith(fetchFn)(DRAFT, { name: 'deck.pptx', bytes: new Uint8Array(3_000_000) });
    expect(answer).toEqual(err({ type: 'network_error', message: `the upload session URL is on ${host}, which is not a host this upload sends a file to` }));
    expect(sent.map((s) => `${s.method} ${s.url}`)).toEqual(['POST https://graph.microsoft.com/v1.0/me/messages/m1/attachments/createUploadSession']);
  });

  it('cancels the upload session, unsigned, and names the byte where Outlook refused a range', async () => {
    const sent: Sent[] = [];
    const fetchFn = graph(sent, (method) => {
      if (method === 'POST') return Response.json({ uploadUrl: UPLOAD_URL }, { status: 201 });
      if (method === 'DELETE') return new Response(null, { status: 204 });
      return sent.filter((s) => s.method === 'PUT').length === 1 ? Response.json({}) : new Response(null, { status: 413 });
    });
    expect(await uploadWith(fetchFn)(DRAFT, { name: 'deck.pptx', bytes: new Uint8Array(4_000_000) })).toEqual(
      err({ type: 'api_error', status: 413, message: 'chunk PUT failed at byte 3276800' })
    );
    expect(sent.at(-1)).toEqual({ method: 'DELETE', url: UPLOAD_URL, authorization: null, type: null, range: null, body: undefined });
  });

  it('answers the error of an upload session Graph refused', async () => {
    const fetchFn = graph([], () => Response.json({ error: { code: 'ErrorItemNotFound', message: 'The draft is gone.' } }, { status: 404 }));
    expect(await uploadWith(fetchFn)(DRAFT, { name: 'deck.pptx', bytes: new Uint8Array(3_000_000) })).toEqual(
      err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound: The draft is gone.', code: 'ErrorItemNotFound' })
    );
  });
});
