import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleEml } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';

const PATH = '/me/messages/m1/attachments/a1';
const EXPANDED = `${PATH}?$expand=microsoft.graph.itemattachment/item`;
// Graph answers a plain read of an embedded Outlook item WITHOUT the item: it
// comes only with $expand, and an embedded mail's own attachments only in $value.
const PLAIN = { '@odata.type': '#microsoft.graph.itemAttachment', id: 'a1', name: 'Legal opinion', contentType: null, size: 15_462_019 };

type Answers = { readonly expanded: Result<unknown, GraphError>; readonly source?: Result<unknown, GraphError> };

const graphWith = (answers: Answers, seen: string[] = []): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    get: async (path: string) => {
      seen.push(path);
      return path === EXPANDED ? answers.expanded : ok(PLAIN);
    },
    getBinary: async (path: string) => {
      seen.push(path);
      return answers.source ?? err({ type: 'api_error', status: 500, message: 'unexpected $value read' });
    },
  });

const EML = new TextDecoder().decode(buildSampleEml());
const MESSAGE: Answers = {
  expanded: ok({ ...PLAIN, item: { '@odata.type': '#microsoft.graph.message', subject: 'Re: Review date — Q3' } }),
  source: ok({ contentType: 'text/plain', size: EML.length, text: EML }),
};

const run = async (name: string, graph: ReturnType<typeof fakeGraphClient>): Promise<Result<unknown, GraphError>> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  return command.execute(graph, { messageId: 'm1', attachmentId: 'a1' });
};

describe('an Outlook item attached to a mail', () => {
  for (const name of ['read-mail-attachment', 'convert-mail-attachment-to-markdown']) {
    it(`${name} reads an embedded mail from its source, its own attachments converted`, async () => {
      const seen: string[] = [];
      const result = await run(name, graphWith(MESSAGE, seen));
      if (!result.ok) throw new Error(result.error.message);
      const text = (result.value as { text: string }).text;
      expect(text).toStartWith('# Re: Review date — Q3');
      expect(text).toContain('### figures.csv');
      expect(text).toContain('| July | 12 |');
      expect(seen).toEqual([PATH, EXPANDED, `${PATH}/$value`]);
    });
  }

  it('renders an embedded meeting from the expanded item, without reading a source', async () => {
    const seen: string[] = [];
    const meeting: Answers = { expanded: ok({ ...PLAIN, item: { '@odata.type': '#microsoft.graph.event', subject: 'Handover walk-through' } }) };
    const result = await run('convert-mail-attachment-to-markdown', graphWith(meeting, seen));
    if (!result.ok) throw new Error(result.error.message);
    expect((result.value as { text: string }).text).toContain('Handover walk-through');
    expect(seen).toEqual([PATH, EXPANDED]);
  });

  it('passes on a failed expansion or a failed source read', async () => {
    const gone = err({ type: 'api_error' as const, status: 404, message: 'ErrorItemNotFound: gone' });
    for (const answers of [{ expanded: gone }, { ...MESSAGE, source: gone }]) {
      const result = await run('read-mail-attachment', graphWith(answers));
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.message).toContain('ErrorItemNotFound: gone');
    }
  });
});
