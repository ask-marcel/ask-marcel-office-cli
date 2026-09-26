import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { executeLocal } from './convert-local-file-to-markdown.ts';
import { commands } from './index.ts';

type Envelope = { readonly contentType?: string; readonly size?: number; readonly text?: string };

const PRD_PAGE = [
  '<html><head><title>PRD</title><style>body { color: red }</style><script>var tracking = 1;</script></head><body>',
  '<header>Team wiki</header><h1>Packaging declaration</h1>',
  '<p>Scope: <a href="https://contoso.example/spec">spec</a></p>',
  '<table><thead><tr><th>Field</th><th>Rule</th></tr></thead><tbody><tr><td>Weight</td><td>kg, 2 decimals</td></tr></tbody></table>',
  '</body></html>',
].join('');

const localPage = async (path: string, html: string): Promise<Envelope> => {
  const fs = createFileSystemFake();
  fs.seedBytes(path, new TextEncoder().encode(html));
  const result = await executeLocal(fs, { path });
  if (!result.ok) throw new Error(result.error.message);
  return result.value as Envelope;
};

describe('an HTML file converted to markdown', () => {
  it('renders the page as markdown, with its head (title, stylesheet, script) dropped and its size counted in UTF-8 bytes', async () => {
    const env = await localPage('/work/prd.html', PRD_PAGE);
    expect(env.contentType).toBe('text/markdown');
    expect(env.text).toStartWith('Team wiki\n\n# Packaging declaration');
    expect(env.text).not.toContain('PRD');
    expect(env.text).toContain('[spec](https://contoso.example/spec)');
    expect(env.text).toContain('| Field | Rule |');
    expect(env.text).toContain('| Weight | kg, 2 decimals |');
    expect(env.text).not.toContain('color: red');
    expect(env.text).not.toContain('tracking');
    expect(env.text).not.toContain('<h1>');
    expect(env.size).toBe(new TextEncoder().encode(env.text).byteLength);
  });

  it('reads an .htm page the same way, a head carrying attributes included', async () => {
    const env = await localPage(
      '/work/status.htm',
      '<html><head profile="https://contoso.example/p"><title>Status</title></head><body><p><strong>Due</strong> Friday — café</p></body></html>'
    );
    expect(env.contentType).toBe('text/markdown');
    expect(env.text).toBe('**Due** Friday — café');
  });

  it('converts an HTML mail attachment through the same dispatch', async () => {
    const attachment = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'upload-requirements.html',
      contentType: 'text/html',
      contentBytes: Buffer.from('<h2>Upload rules</h2><ul><li>One PO per file</li></ul>').toString('base64'),
    };
    const command = commands['read-mail-attachment'];
    if (!command) throw new Error('read-mail-attachment is not registered');
    const result = await command.execute(fakeGraphClient({ get: async () => ok(attachment) }), { messageId: 'm1', attachmentId: 'a1' });
    if (!result.ok) throw new Error(result.error.message);
    const env = result.value as Envelope;
    expect(env.contentType).toBe('text/markdown');
    expect(env.text).toContain('## Upload rules');
    expect(env.text).toMatch(/-\s+One PO per file/);
  });
});
