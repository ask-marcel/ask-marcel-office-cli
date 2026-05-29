import JSZip from 'jszip';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Thin wrapper around JSZip for opening already-fetched OOXML bytes
 * (.docx / .xlsx / .pptx — all ZIP packages with a shared docProps/* core).
 * Pure CPU on bytes already in memory — no IO — so this matches the
 * mammoth-adapter pattern: a Result-returning factory, no port, no fake.
 * Tests work end-to-end against real fixtures built with the `docx` /
 * `xlsx` packages (or hand-rolled JSZip) via test-helpers/office-fixtures.ts.
 *
 * try/catch is permitted under src/infra/** (atelier hard rule 17): any
 * malformed-zip throw from JSZip translates into a Result.err with a
 * GraphError shape so callers stay on the Result rail.
 *
 * The returned reader is synchronous on purpose: callers in the use-case
 * layer are pure XML walkers that don't model async/await internally.
 * Every text entry is pre-decoded once at open time, then served from
 * an in-memory map. `list()` exposes the entry paths so callers can
 * enumerate numbered parts (`xl/comments1.xml`, `ppt/slides/slide*.xml`,
 * every `*.rels`) that can't be hardcoded.
 */
type OoxmlZip = {
  readonly read: (path: string) => string | undefined;
  readonly list: () => ReadonlyArray<string>;
};

const openOoxmlZip = async (bytes: Uint8Array): Promise<Result<OoxmlZip, GraphError>> => {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const entries = Object.keys(zip.files)
      .map((name) => ({ name, file: zip.file(name) }))
      .filter((e): e is { name: string; file: JSZip.JSZipObject } => e.file !== null && !e.file.dir);
    const contents = await Promise.all(entries.map((e) => e.file.async('string')));
    const memo = new Map<string, string>();
    entries.forEach((e, i) => memo.set(e.name, contents[i] ?? ''));
    const read = (path: string): string | undefined => memo.get(path);
    const list = (): ReadonlyArray<string> => [...memo.keys()];
    return ok({ read, list });
  } catch (e) {
    return err({ type: 'api_error', status: 500, message: `ooxml zip parse failed: ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { openOoxmlZip };
export type { OoxmlZip };
