import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { attrOf, collectText, findAll, parseXml, textOf } from './ooxml-xml-walker.ts';

/**
 * The OOXML package-level metadata that is identical across .docx / .xlsx /
 * .pptx: the three `docProps/*` property parts and every external
 * relationship in the package. Per-format modules (docx-metadata,
 * xlsx-metadata, …) compose these with their own body-specific extractors.
 */

type CustomProp = { readonly name: string; readonly value: string };
type ExternalRel = { readonly source: string; readonly type: string; readonly target: string };

const stripNs = (tag: string): string => {
  const colon = tag.indexOf(':');
  return colon === -1 ? tag : tag.slice(colon + 1);
};

// `cp:coreProperties` / `Properties` (extended) — children are direct
// text-only elements like `<dc:creator>Vincent</dc:creator>`. Flatten to
// a Record by stripping the namespace prefix from the local name.
const flattenLeafProps = (root: unknown, rootTag: string): Readonly<Record<string, string>> => {
  const containers = findAll(root, rootTag);
  if (containers.length === 0) return {};
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(containers[0] ?? {})) {
    if (key.startsWith('@_') || key === '#text') continue;
    // app.xml has nested vt:lpstr / vt:variant for some keys — flatten them too
    const text = typeof value === 'string' ? value : textOf(value) || collectText(value, 'vt:lpstr') || collectText(value, 'vt:variant');
    if (text !== '') out[stripNs(key)] = text;
  }
  return out;
};

const extractCoreProps = (zip: OoxmlZip): Readonly<Record<string, string>> => flattenLeafProps(parseXml(zip.read('docProps/core.xml')), 'cp:coreProperties');

const extractAppProps = (zip: OoxmlZip): Readonly<Record<string, string>> => flattenLeafProps(parseXml(zip.read('docProps/app.xml')), 'Properties');

const extractCustomProps = (zip: OoxmlZip): ReadonlyArray<CustomProp> => {
  const props = findAll(parseXml(zip.read('docProps/custom.xml')), 'property');
  const out: Array<CustomProp> = [];
  for (const p of props) {
    const name = attrOf(p, 'name');
    let value = '';
    for (const [k, v] of Object.entries(p)) {
      if (k.startsWith('@_') || k === '#text') continue;
      value = textOf(v);
      if (value !== '') break;
    }
    if (name !== '') out.push({ name, value });
  }
  return out;
};

const lastSegment = (s: string): string => {
  const slash = s.lastIndexOf('/');
  return slash === -1 ? s : s.slice(slash + 1);
};

const extractRelsFromOne = (root: unknown, source: string): ReadonlyArray<ExternalRel> => {
  const out: Array<ExternalRel> = [];
  for (const r of findAll(root, 'Relationship')) {
    if (attrOf(r, 'TargetMode') !== 'External') continue;
    out.push({ source, type: lastSegment(attrOf(r, 'Type')), target: attrOf(r, 'Target') });
  }
  return out;
};

// Scan every `*.rels` part in the package — the relationship graph is split
// across many parts (`_rels/.rels`, `word/_rels/document.xml.rels`,
// `xl/externalLinks/_rels/*.rels`, …) and the numbered ones can't be
// hardcoded. Only `TargetMode="External"` targets are surfaced.
const extractExternalRels = (zip: OoxmlZip): ReadonlyArray<ExternalRel> => {
  const out: Array<ExternalRel> = [];
  for (const path of zip.list()) {
    if (!path.endsWith('.rels')) continue;
    for (const rel of extractRelsFromOne(parseXml(zip.read(path)), path)) out.push(rel);
  }
  return out;
};

export { extractAppProps, extractCoreProps, extractCustomProps, extractExternalRels };
export type { CustomProp, ExternalRel };
