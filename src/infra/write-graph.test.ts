import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import { createReadGraph } from './read-graph.ts';
import { createWriteGraph, writeGraphOf } from './write-graph.ts';

/*
 * The write graph: basic tier only, the members the write commands use. A
 * write that succeeds with an empty body (202 Accepted, 204 No Content) is a
 * success: reading it as JSON would report a network error, and the caller
 * might send the write again.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

// A Graph that answers every request with a fresh copy of `response`.
const answering = (response: Response): FetchFn => {
  const fetchFn: FetchFn = async () => response.clone();
  return fetchFn;
};

const writeGraphAnswering = (response: Response): ReturnType<typeof createWriteGraph> => createWriteGraph(tokens, answering(response));

const WRITE_GRAPH_MEMBERS = ['delete', 'fetchUrl', 'get', 'getBinary', 'patch', 'post', 'put'];

describe('the write graph', () => {
  it('answers a POST that Graph accepted with no body (202) as a success with no value', async () => {
    expect(await writeGraphAnswering(new Response(null, { status: 202 })).post('/me/messages/m1/send', {})).toEqual(ok(undefined));
  });

  it('answers a PATCH that Graph applied with no content (204) as a success with no value', async () => {
    expect(await writeGraphAnswering(new Response(null, { status: 204 })).patch('/me/messages/m1', { subject: 'Hi' })).toEqual(ok(undefined));
  });

  it('still answers a write with the JSON Graph sends back', async () => {
    expect(await writeGraphAnswering(Response.json({ id: 'draft-1' }, { status: 201 })).post('/me/messages', { subject: 'Hi' })).toEqual(ok({ id: 'draft-1' }));
  });

  it('reports a body that is not JSON as a network error, as before', async () => {
    const result = await writeGraphAnswering(new Response('<html>', { status: 200 })).post('/me/messages', {});
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('network_error');
  });

  it('holds only the basic-tier members the write commands use', () => {
    expect(Object.keys(writeGraphAnswering(Response.json({}))).toSorted((a, b) => a.localeCompare(b))).toEqual(WRITE_GRAPH_MEMBERS);
  });
});

describe('the write view of a graph that can also read every tier', () => {
  it('keeps only the write graph members, so a write command never holds an elevated, guest or chat reader', () => {
    const full = { ...createReadGraph(tokens, answering(Response.json({}))), ...writeGraphAnswering(Response.json({})) };
    expect(Object.keys(writeGraphOf(full)).toSorted((a, b) => a.localeCompare(b))).toEqual(WRITE_GRAPH_MEMBERS);
  });

  it("sends each write and read to the graph's own member of that name, on a caller's graph whose members are methods that read `this`", async () => {
    const graph = {
      calls: [] as string[],
      async get(path: string) {
        this.calls.push(`get ${path}`);
        return ok({});
      },
      async getBinary(path: string) {
        this.calls.push(`getBinary ${path}`);
        return ok({});
      },
      async fetchUrl(url: string) {
        this.calls.push(`fetchUrl ${url}`);
        return ok({});
      },
      async post(path: string) {
        this.calls.push(`post ${path}`);
        return ok({});
      },
      async patch(path: string) {
        this.calls.push(`patch ${path}`);
        return ok({});
      },
      async put(path: string) {
        this.calls.push(`put ${path}`);
        return ok({});
      },
      async delete(path: string) {
        this.calls.push(`delete ${path}`);
        return ok(undefined);
      },
    };
    const view = writeGraphOf(graph);
    await view.get('/1');
    await view.getBinary('/2');
    await view.fetchUrl('https://3');
    await view.post('/4', {});
    await view.patch('/5', {});
    await view.put('/6', new Uint8Array(1));
    await view.delete('/7');
    expect(graph.calls).toEqual(['get /1', 'getBinary /2', 'fetchUrl https://3', 'post /4', 'patch /5', 'put /6', 'delete /7']);
  });
});

describe('a file upload Graph answers with no body', () => {
  it('answers a small upload that Graph stored with no content (204) as a success with no value', async () => {
    expect(await writeGraphAnswering(new Response(null, { status: 204 })).put('/me/drive/root:/.ask-marcel-temp/a.docx', new Uint8Array(1))).toEqual(ok(undefined));
  });

  it('reports a large upload whose upload session came back empty as a Graph answer with no upload URL, not a network error', async () => {
    const result = await writeGraphAnswering(new Response(null, { status: 200 })).put('/me/drive/root:/.ask-marcel-temp/a.docx', new Uint8Array(4 * 1024 * 1024 + 1));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toEqual({ type: 'api_error', status: 500, message: 'createUploadSession returned no uploadUrl' });
  });
});

describe('what the write graph sends', () => {
  type Sent = { readonly method: string; readonly url: string; readonly prefer: string | null };
  const recording = (sent: Sent[]): FetchFn => {
    const fetchFn: FetchFn = async (url, init) => {
      sent.push({ method: init?.method ?? '', url, prefer: new Headers(init?.headers).get('prefer') });
      return Response.json({ id: 'item-1' });
    };
    return fetchFn;
  };

  it('passes the headers its caller adds on a read', async () => {
    const sent: Sent[] = [];
    await createWriteGraph(tokens, recording(sent)).get('/me/messages/delta', { Prefer: 'odata.maxpagesize=5' });
    expect(sent).toEqual([{ method: 'GET', url: 'https://graph.microsoft.com/v1.0/me/messages/delta', prefer: 'odata.maxpagesize=5' }]);
  });

  it('uploads a file of exactly 4 MiB in one PUT, and a larger one through an upload session', async () => {
    const sent: Sent[] = [];
    await createWriteGraph(tokens, recording(sent)).put('/me/drive/root:/.ask-marcel-temp/a.docx', new Uint8Array(4 * 1024 * 1024));
    expect(sent.map((s) => `${s.method} ${s.url}`)).toEqual(['PUT https://graph.microsoft.com/v1.0/me/drive/root:/.ask-marcel-temp/a.docx:/content']);
    const larger: Sent[] = [];
    await createWriteGraph(tokens, recording(larger)).put('/me/drive/root:/.ask-marcel-temp/a.docx', new Uint8Array(4 * 1024 * 1024 + 1));
    expect(larger[0]).toEqual({ method: 'POST', url: 'https://graph.microsoft.com/v1.0/me/drive/root:/.ask-marcel-temp/a.docx:/createUploadSession', prefer: null });
  });
});

describe('how the write graph uploads a file', () => {
  type Upload = { readonly method: string; readonly contentType: string | null; readonly body: unknown };
  const recordingUploads = (sent: Upload[]): FetchFn => {
    const fetchFn: FetchFn = async (_url, init) => {
      sent.push({ method: init?.method ?? '', contentType: new Headers(init?.headers).get('content-type'), body: init?.body });
      return Response.json({ id: 'item-1' });
    };
    return fetchFn;
  };

  it('labels a small upload with the content type its caller names, and as plain bytes when it names none', async () => {
    const sent: Upload[] = [];
    const graph = createWriteGraph(tokens, recordingUploads(sent));
    await graph.put('/me/drive/root:/.ask-marcel-temp/a.docx', new Uint8Array(1), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    await graph.put('/me/drive/root:/.ask-marcel-temp/b.bin', new Uint8Array(1));
    expect(sent.map((upload) => upload.contentType)).toEqual(['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/octet-stream']);
  });

  it('opens an upload session that replaces a temporary file of the same name', async () => {
    const sent: Upload[] = [];
    await createWriteGraph(tokens, recordingUploads(sent)).put('/me/drive/root:/.ask-marcel-temp/a.docx', new Uint8Array(4 * 1024 * 1024 + 1));
    expect(sent[0]?.method).toBe('POST');
    expect(JSON.parse(String(sent[0]?.body))).toEqual({ item: { '@microsoft.graph.conflictBehavior': 'replace' } });
  });

  it('names the upload and the file-transfer deadline when a small upload times out', async () => {
    const timingOut: FetchFn = async () => {
      throw new DOMException('timed out', 'TimeoutError');
    };
    const result = await createWriteGraph(tokens, timingOut).put('/me/drive/root:/.ask-marcel-temp/a.docx', new Uint8Array(1));
    expect(result.ok).toBe(false);
    if (!result.ok)
      expect(result.error.message).toBe('request timed out after 5min (PUT /me/drive/root:/.ask-marcel-temp/a.docx:/content) — transient; retry once before treating as permanent');
  });
});
