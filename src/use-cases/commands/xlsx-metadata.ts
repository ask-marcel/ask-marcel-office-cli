import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractAppProps, extractCoreProps, extractCustomProps, extractExternalRels, extractMacros } from './ooxml-metadata.ts';
import type { CustomProp, ExternalRel } from './ooxml-metadata.ts';
import { attrOf, findAll, parseXml, textOf } from './ooxml-xml-walker.ts';
import { extractLegacyComments, extractPeople, extractThreadedComments } from './xlsx-comments.ts';
import type { CellComment, Person, ThreadedComment } from './xlsx-comments.ts';

/**
 * Pulls the side-channel content out of a .xlsx zip — everything a user can
 * author that the value-rendered markdown body (cells per sheet) never shows:
 * core / app / custom doc properties, external relationships, defined names
 * (named ranges & formulas), hidden / very-hidden sheets, legacy cell
 * comments, threaded comments, and the persons registry behind them.
 *
 * Package-level parts (docProps/*, every *.rels) come from the shared
 * ooxml-metadata module; this file owns the workbook-specific parts.
 */

type PropMap = Readonly<Record<string, string>>;
type DefinedName = { readonly name: string; readonly refersTo: string; readonly hidden: boolean };
type Sheet = { readonly name: string; readonly state: string };

type XlsxMetadata = {
  readonly core: PropMap;
  readonly app: PropMap;
  readonly custom: ReadonlyArray<CustomProp>;
  readonly externalRels: ReadonlyArray<ExternalRel>;
  readonly definedNames: ReadonlyArray<DefinedName>;
  readonly hiddenSheets: ReadonlyArray<Sheet>;
  readonly comments: ReadonlyArray<CellComment>;
  readonly threadedComments: ReadonlyArray<ThreadedComment>;
  readonly people: ReadonlyArray<Person>;
  readonly macros: ReadonlyArray<string>;
};

const HIDDEN_STATES: ReadonlySet<string> = new Set(['hidden', 'veryHidden']);

const extractDefinedNames = (workbook: unknown): ReadonlyArray<DefinedName> =>
  findAll(workbook, 'definedName').map((dn) => ({
    name: attrOf(dn, 'name'),
    refersTo: textOf(dn),
    hidden: attrOf(dn, 'hidden') === '1' || attrOf(dn, 'hidden') === 'true',
  }));

const extractHiddenSheets = (workbook: unknown): ReadonlyArray<Sheet> =>
  findAll(workbook, 'sheet')
    .map((s) => ({ name: attrOf(s, 'name'), state: attrOf(s, 'state') }))
    .filter((s) => HIDDEN_STATES.has(s.state));

const extractXlsxMetadata = async (bytes: Uint8Array): Promise<Result<XlsxMetadata, GraphError>> => {
  const zipR = await openOoxmlZip(bytes);
  if (!zipR.ok) return zipR;
  const zip = zipR.value;
  const workbook = parseXml(zip.read('xl/workbook.xml'));
  const people = extractPeople(zip);
  return ok({
    core: extractCoreProps(zip),
    app: extractAppProps(zip),
    custom: extractCustomProps(zip),
    externalRels: extractExternalRels(zip),
    definedNames: extractDefinedNames(workbook),
    hiddenSheets: extractHiddenSheets(workbook),
    comments: extractLegacyComments(zip),
    threadedComments: extractThreadedComments(zip, people),
    people,
    macros: extractMacros(zip),
  });
};

export { extractXlsxMetadata };
export type { CellComment, DefinedName, Person, Sheet, ThreadedComment, XlsxMetadata };
