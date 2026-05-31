import { XMLParser } from 'fast-xml-parser';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { escapeCell } from './ooxml-metadata-to-markdown.ts';

/**
 * Converts an OpenDocument body (`content.xml`) to markdown. ODF stores its
 * body in document order under `office:body > {office:text | office:spreadsheet
 * | office:presentation}`, with all-different namespace prefixes (text: / table:
 * / draw:) from OOXML. The shared walker drops cross-element order (fast-xml-parser
 * groups same-named siblings), so this module parses with `preserveOrder` and walks
 * the ordered tree itself: headings, paragraphs, lists, and tables for text docs;
 * named sheet-tables for spreadsheets; per-slide text for presentations.
 *
 * Nothing is hidden: walking the raw content.xml surfaces every text run regardless
 * of the style-driven visibility a rendered viewer would apply, and a
 * `text:section` flagged `text:display="none"` is emitted with an explicit marker
 * rather than dropped — the whole point of feeding this to an LLM.
 *
 * Pure (string → string); the only IO is `openOoxmlZip`, which returns a Result.
 */

type Node = Record<string, unknown>;

const TEXT = '#text';
const ATTRS = ':@';
// ODS pads its trailing empty cells/rows with huge `number-*-repeated` counts
// (often 16384); cap the expansion so a blank tail can't blow up, then trim it.
const MAX_REPEAT = 1024;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', parseAttributeValue: false, parseTagValue: false, trimValues: false, preserveOrder: true });

const tagOf = (node: Node): string | undefined => {
  for (const key of Object.keys(node)) if (key !== ATTRS && key !== TEXT) return key;
  return undefined;
};

const kidsOf = (node: Node, tag: string): ReadonlyArray<Node> => {
  const value = node[tag];
  return Array.isArray(value) ? (value as ReadonlyArray<Node>) : [];
};

const attr = (node: Node, name: string): string => {
  const bag = node[ATTRS];
  if (bag === null || typeof bag !== 'object') return '';
  const value = (bag as Record<string, unknown>)[`@_${name}`];
  return typeof value === 'string' ? value : '';
};

const intAttr = (node: Node, name: string): number => {
  const parsed = Number.parseInt(attr(node, name), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 1;
};

const inlineText = (children: ReadonlyArray<Node>): string => {
  let out = '';
  for (const node of children) {
    const leaf = node[TEXT];
    if (typeof leaf === 'string') {
      out += leaf;
      continue;
    }
    const tag = tagOf(node);
    if (tag === undefined) continue;
    if (tag === 'text:s') out += ' '.repeat(intAttr(node, 'text:c'));
    else if (tag === 'text:tab' || tag === 'text:line-break') out += ' ';
    else out += inlineText(kidsOf(node, tag));
  }
  return out;
};

const renderHeading = (node: Node, kids: ReadonlyArray<Node>): string => {
  const level = Math.min(intAttr(node, 'text:outline-level'), 6);
  return `${'#'.repeat(level)} ${inlineText(kids).trim()}`;
};

const renderListLines = (listKids: ReadonlyArray<Node>, depth: number): ReadonlyArray<string> => {
  const out: Array<string> = [];
  const pad = '  '.repeat(depth);
  for (const item of listKids) {
    const itemTag = tagOf(item);
    if (itemTag !== 'text:list-item' && itemTag !== 'text:list-header') continue;
    for (const child of kidsOf(item, itemTag)) {
      const childTag = tagOf(child);
      if (childTag === undefined) continue;
      if (childTag === 'text:list') out.push(...renderListLines(kidsOf(child, childTag), depth + 1));
      else {
        const text = inlineText(kidsOf(child, childTag)).trim();
        if (text !== '') out.push(`${pad}- ${text}`);
      }
    }
  }
  return out;
};

const cellText = (cell: Node, tag: string): string =>
  kidsOf(cell, tag)
    .filter((child) => tagOf(child) === 'text:p' || tagOf(child) === 'text:h')
    .map((para) => inlineText(kidsOf(para, tagOf(para) ?? '')).trim())
    .filter((text) => text !== '')
    .join(' ');

const trimTrailingEmpty = <T>(items: ReadonlyArray<T>, isEmpty: (item: T) => boolean): ReadonlyArray<T> => {
  let end = items.length;
  while (end > 0 && isEmpty(items[end - 1] as T)) end -= 1;
  return items.slice(0, end);
};

const rowCells = (row: Node): ReadonlyArray<string> => {
  const cells: Array<string> = [];
  for (const cell of kidsOf(row, 'table:table-row')) {
    const cellTag = tagOf(cell);
    if (cellTag !== 'table:table-cell' && cellTag !== 'table:covered-table-cell') continue;
    const text = cellText(cell, cellTag);
    for (let i = 0; i < Math.min(intAttr(cell, 'table:number-columns-repeated'), MAX_REPEAT); i += 1) cells.push(text);
  }
  return trimTrailingEmpty(cells, (text) => text === '');
};

const tableRows = (tableKids: ReadonlyArray<Node>): ReadonlyArray<ReadonlyArray<string>> => {
  const rows: Array<ReadonlyArray<string>> = [];
  for (const node of tableKids) {
    const tag = tagOf(node);
    if (tag === 'table:table-header-rows') {
      rows.push(...tableRows(kidsOf(node, tag)));
      continue;
    }
    if (tag !== 'table:table-row') continue;
    const cells = rowCells(node);
    for (let i = 0; i < Math.min(intAttr(node, 'table:number-rows-repeated'), MAX_REPEAT); i += 1) rows.push(cells);
  }
  return rows;
};

const renderTable = (tableKids: ReadonlyArray<Node>): string => {
  const rows = trimTrailingEmpty(tableRows(tableKids), (row) => row.every((cell) => cell === ''));
  if (rows.length === 0) return '';
  const colCount = Math.max(...rows.map((row) => row.length), 1);
  const padRow = (row: ReadonlyArray<string>): string => `| ${Array.from({ length: colCount }, (_unused, i) => escapeCell(row[i] ?? '')).join(' | ')} |`;
  const separator = `| ${Array.from({ length: colCount }, () => '---').join(' | ')} |`;
  return [padRow(rows[0] ?? []), separator, ...rows.slice(1).map(padRow)].join('\n');
};

const renderBlock = (node: Node, tag: string): ReadonlyArray<string> => {
  const kids = kidsOf(node, tag);
  if (tag === 'text:h') return [renderHeading(node, kids)];
  if (tag === 'text:p') {
    const text = inlineText(kids).trim();
    return text === '' ? [] : [text];
  }
  if (tag === 'text:list') {
    const lines = renderListLines(kids, 0);
    return lines.length === 0 ? [] : [lines.join('\n')];
  }
  if (tag === 'table:table') {
    const table = renderTable(kids);
    return table === '' ? [] : [table];
  }
  if (tag === 'text:section') {
    const blocks = renderBlocks(kids);
    if (blocks.length === 0) return [];
    return attr(node, 'text:display') === 'none' ? ['> _(hidden section — not shown in the rendered document)_', ...blocks] : blocks;
  }
  return renderBlocks(kids);
};

const renderBlocks = (children: ReadonlyArray<Node>): ReadonlyArray<string> => {
  const out: Array<string> = [];
  for (const node of children) {
    const tag = tagOf(node);
    if (tag !== undefined) out.push(...renderBlock(node, tag));
  }
  return out;
};

const renderSheets = (kids: ReadonlyArray<Node>): ReadonlyArray<string> => {
  const out: Array<string> = [];
  for (const node of kids) {
    if (tagOf(node) !== 'table:table') continue;
    const name = attr(node, 'table:name');
    out.push(name === '' ? '## Sheet' : `## ${name}`);
    const table = renderTable(kidsOf(node, 'table:table'));
    if (table !== '') out.push(table);
  }
  return out;
};

const renderSlides = (kids: ReadonlyArray<Node>): ReadonlyArray<string> => {
  const out: Array<string> = [];
  let index = 0;
  for (const node of kids) {
    if (tagOf(node) !== 'draw:page') continue;
    index += 1;
    const name = attr(node, 'draw:name');
    const label = name === '' ? `Slide ${index}` : name;
    out.push(`## ${label}`);
    out.push(...renderBlocks(kidsOf(node, 'draw:page')));
  }
  return out;
};

const findFirst = (nodes: ReadonlyArray<Node>, tag: string): Node | undefined => {
  for (const node of nodes) {
    const nodeTag = tagOf(node);
    if (nodeTag === undefined) continue;
    if (nodeTag === tag) return node;
    const found = findFirst(kidsOf(node, nodeTag), tag);
    if (found !== undefined) return found;
  }
  return undefined;
};

const renderBody = (bodyKids: ReadonlyArray<Node>): ReadonlyArray<string> => {
  const root = bodyKids.find((node) => tagOf(node) !== undefined);
  if (root === undefined) return [];
  const tag = tagOf(root) ?? '';
  const kids = kidsOf(root, tag);
  if (tag === 'office:spreadsheet') return renderSheets(kids);
  if (tag === 'office:presentation') return renderSlides(kids);
  return renderBlocks(kids);
};

const renderOdfContent = (xml: string | undefined): string => {
  if (xml === undefined || xml === '') return '';
  const body = findFirst(parser.parse(xml) as ReadonlyArray<Node>, 'office:body');
  if (body === undefined) return '';
  return renderBody(kidsOf(body, 'office:body')).join('\n\n');
};

const odfContentToMarkdown = async (bytes: Uint8Array): Promise<Result<string, GraphError>> => {
  const zipR = await openOoxmlZip(bytes);
  if (!zipR.ok) return zipR;
  return ok(renderOdfContent(zipR.value.read('content.xml')));
};

export { odfContentToMarkdown, renderOdfContent };
