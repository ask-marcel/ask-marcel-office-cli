import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { attrOf, findAll, findAllTexts, parseXml, textOf } from './ooxml-xml-walker.ts';

/**
 * Pulls the side-channel content out of an OpenDocument package (.odt / .ods /
 * .odp and their .ot* template variants). OpenDocument is also a ZIP, so the
 * shared `openOoxmlZip` adapter + XML walker apply directly; only the metadata
 * shape differs — it lives in a single `meta.xml` (`office:document-meta >
 * office:meta`) rather than the OOXML `docProps/*` parts.
 *
 * High-value subset: document properties (Dublin Core + ODF meta fields like
 * generator / editing-cycles / creation-date), keyword list, and user-defined
 * custom properties (`meta:user-defined`) — the ODF analog of OOXML custom
 * document properties.
 */

type UserDefined = { readonly name: string; readonly value: string };
type OdfMetadata = {
  readonly properties: Readonly<Record<string, string>>;
  readonly keywords: ReadonlyArray<string>;
  readonly userDefined: ReadonlyArray<UserDefined>;
};

// Tags handled by their own section, so they are not folded into the flat
// property record.
const SPECIAL_TAGS: ReadonlySet<string> = new Set(['meta:keyword', 'meta:user-defined']);

const stripNs = (tag: string): string => {
  const colon = tag.indexOf(':');
  return colon === -1 ? tag : tag.slice(colon + 1);
};

const extractProperties = (meta: unknown): Readonly<Record<string, string>> => {
  const container = findAll(meta, 'office:meta')[0];
  if (container === undefined) return {};
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(container)) {
    if (key.startsWith('@_') || key === '#text' || SPECIAL_TAGS.has(key)) continue;
    const text = textOf(value);
    if (text !== '') out[stripNs(key)] = text;
  }
  return out;
};

const extractUserDefined = (meta: unknown): ReadonlyArray<UserDefined> =>
  findAll(meta, 'meta:user-defined')
    .map((p) => ({ name: attrOf(p, 'meta:name'), value: textOf(p) }))
    .filter((p) => p.name !== '');

const extractOdfMetadata = async (bytes: Uint8Array): Promise<Result<OdfMetadata, GraphError>> => {
  const zipR = await openOoxmlZip(bytes);
  if (!zipR.ok) return zipR;
  const meta = parseXml(zipR.value.read('meta.xml'));
  return ok({
    properties: extractProperties(meta),
    keywords: findAllTexts(meta, 'meta:keyword').filter((k) => k.trim() !== ''),
    userDefined: extractUserDefined(meta),
  });
};

export { extractOdfMetadata };
export type { OdfMetadata, UserDefined };
