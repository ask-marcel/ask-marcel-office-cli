import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['download-drive-item-version'];
const ARGS = { driveId: 'b!x', itemId: '01ABC', versionId: '2.0' };

// Graph answers `?format=html` on a historical version with the CURRENT page
// (probed 2026-09-26: identical HTML for every version whose raw bytes differ),
// so a markdown rendering of an old Loop version would be the wrong page.
describe('a historical version of a Loop or Whiteboard file', () => {
  it('is refused in markdown, naming the two reads that do work, before asking Graph to render anything', async () => {
    if (!command) throw new Error('download-drive-item-version is not registered');
    for (const name of ['notes.loop', 'board.whiteboard', 'draft.fluid', 'sketch.wbtx']) {
      const calls: string[] = [];
      const graph = fakeGraphClient({
        get: async (path) => {
          calls.push(path);
          return ok({ name });
        },
        getBinaryElevated: async (path) => {
          calls.push(path);
          return ok({ contentType: 'text/html', size: 4, base64: 'PHA+' });
        },
      });
      const r = await command.execute(graph, { ...ARGS, format: 'markdown' });
      expect(r.ok).toBe(false);
      if (!r.ok) {
        expect(r.error).toMatchObject({ type: 'api_error', status: 415, code: 'unsupported_version_render' });
        expect(r.error.message).toContain('--format original');
        expect(r.error.message).toContain('download-drive-item-as-markdown');
      }
      expect(calls.some((path) => path.includes('format=html'))).toBe(false);
    }
  });
});
