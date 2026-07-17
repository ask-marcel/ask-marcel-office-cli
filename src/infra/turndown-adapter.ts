import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';
import type { Result } from '../domain/result.ts';
import { ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Turn arbitrary HTML — mammoth-converted docx, Graph-converted Office
 * docs, OneNote pages, Outlook email bodies, SharePoint pages — into
 * clean markdown.
 *
 * Wrapped in a function-only adapter so the rest of the codebase stays
 * class-free per atelier rule 1. Each call constructs a fresh service
 * (turndown is cheap to instantiate; no mutable shared state).
 *
 * GFM plugin enabled by default so HTML tables render as pipe-delimited
 * markdown tables instead of flat paragraphs (needed for docx → markdown
 * to preserve table structure that mammoth emits as <table><tr><td>).
 *
 * Turndown's default HTML parser drops comment nodes during parsing,
 * so we don't need a custom comment-stripping rule. <script> and
 * <style> aren't dropped by default — those need explicit removal so
 * their text content doesn't leak into the markdown output.
 *
 * Lives in `src/infra/` because turndown can throw on malformed DOM. The
 * dominant culprit is the GFM table plugin's `cellInternals` walker,
 * which assumes every `<td>` has an ancestor `<tr>` reachable via
 * `parentNode` — Outlook MSO HTML routinely violates this with floating
 * `<td>` cells, nested misnested tables, or `<table>` wrappers with no
 * `<tbody>`. The walker then dereferences `undefined.parentNode` and
 * throws `Cannot read properties of undefined (reading 'parentNode')`.
 *
 * Three-tier graceful degradation when that happens:
 *
 *   1. turndown WITH GFM (clean markdown tables)
 *   2. turndown WITHOUT GFM (core markdown; tables become flat paragraphs
 *      but the rest of the document converts cleanly)
 *   3. stripped-text fallback (state-machine HTML walker; preserves text,
 *      basic entities, and newlines on block boundaries)
 *
 * Tier 2 covers the vast majority of Outlook MSO bodies; tier 3 only
 * fires when even the core walker chokes. Each downgrade prepends a
 * markdown blockquote note naming the underlying error so observability
 * survives the degradation.
 *
 * Headerless tables are a second, silent failure mode of the GFM plugin:
 * its `keep()` rule passes any table without a proper `<th>` heading row
 * through as RAW `outerHTML` (a pasted-from-Excel MsoNormalTable emitted
 * 140 KB of styled markup for ~3 KB of cell text). The `headerlessTable`
 * rule below intercepts exactly those tables (addRule'd rules are checked
 * before keep rules) and degrades them to a pipe table, or unwraps them
 * when they are single-column Outlook layout scaffolding. It also absorbs
 * the empty `<table></table>` case whose `rows[0]` dereference used to
 * throw inside the keep filter and drag the whole document to tier 2.
 */
type DomNode = {
  readonly nodeName: string;
  readonly childNodes: ArrayLike<DomNode>;
  readonly parentNode: DomNode | null;
  readonly firstChild: DomNode | null;
  readonly previousSibling: DomNode | null;
  readonly textContent: string;
  readonly rows: ArrayLike<DomNode>;
  readonly getAttribute: (name: string) => string | null;
};

// Replicas of turndown-plugin-gfm's isFirstTbody/isHeadingRow, so the rule
// filter below is the exact complement of the plugin's own `table` rule:
// properly-headed tables keep flowing to the plugin, everything else is ours.
const isFirstTbodyLike = (el: DomNode): boolean =>
  el.nodeName === 'TBODY' && (el.previousSibling === null || (el.previousSibling.nodeName === 'THEAD' && /^\s*$/.test(el.previousSibling.textContent)));

const isHeadingRowLike = (tr: DomNode): boolean => {
  const parent = tr.parentNode;
  if (parent === null) return false;
  if (parent.nodeName === 'THEAD') return true;
  return parent.firstChild === tr && (parent.nodeName === 'TABLE' || isFirstTbodyLike(parent)) && Array.from(tr.childNodes).every((n) => n.nodeName === 'TH');
};

const hasProperHeader = (table: DomNode): boolean => table.rows.length > 0 && isHeadingRowLike(table.rows[0]);

// Turndown's DOM (domino) implements `.rows` / `.cells` as DEEP
// getElementsByTagName walks that include rows of tables nested inside
// cells, so the serializer walks childNodes shallowly instead.
const SECTION_TAGS: ReadonlySet<string> = new Set(['THEAD', 'TBODY', 'TFOOT']);

const childrenOf = (node: DomNode): ReadonlyArray<DomNode> => Array.from(node.childNodes);

const shallowRows = (table: DomNode): ReadonlyArray<DomNode> =>
  childrenOf(table).flatMap((child) => {
    if (child.nodeName === 'TR') return [child];
    if (SECTION_TAGS.has(child.nodeName)) return childrenOf(child).filter((n) => n.nodeName === 'TR');
    return [];
  });

const shallowCells = (tr: DomNode): ReadonlyArray<DomNode> => childrenOf(tr).filter((n) => n.nodeName === 'TD' || n.nodeName === 'TH');

// Bounds best-effort colspan padding on hostile input (colspan="9999").
const COLSPAN_CAP = 20;

const spanOf = (cell: DomNode): number => {
  const parsed = Number.parseInt(cell.getAttribute('colspan') ?? '1', 10);
  return Math.min(Math.max(Number.isNaN(parsed) ? 1 : parsed, 1), COLSPAN_CAP);
};

const addHeaderlessTableRule = (td: TurndownService): void => {
  // Reentrant `td.turndown(cell)` keeps inline formatting (bold, links,
  // entities) exactly like the plugin's own cell handling; turndown clones
  // element inputs and holds no per-run instance state, and each reentrant
  // call converts a strict subtree, so the recursion terminates. Pipe cells
  // are single-line, and turndown never escapes `|` itself.
  const convertCell = (cell: DomNode): string =>
    td
      .turndown(cell as unknown as HTMLElement)
      .replaceAll(/\s+/g, ' ')
      .replaceAll('|', '\\|')
      .trim();
  td.addRule('headerlessTable', {
    filter: (node) => node.nodeName === 'TABLE' && !hasProperHeader(node as unknown as DomNode),
    replacement: (_content, node) => {
      const rows = shallowRows(node as unknown as DomNode);
      if (rows.length === 0) return '';
      const cellsPerRow = rows.map(shallowCells);
      const width = Math.max(...cellsPerRow.map((cells) => cells.reduce((n, c) => n + spanOf(c), 0)));
      if (width <= 1) {
        // Single-cell / single-column tables are Outlook layout scaffolding,
        // not data: unwrap each cell as free-standing blocks. Nested tables
        // recurse back through this rule.
        const blocks = cellsPerRow
          .flat()
          .map((cell) => td.turndown(cell as unknown as HTMLElement).trim())
          .filter((s) => s.length > 0);
        return blocks.length === 0 ? '' : `\n\n${blocks.join('\n\n')}\n\n`;
      }
      const grid = cellsPerRow.map((cells) => {
        const out: string[] = [];
        for (const cell of cells) {
          out.push(convertCell(cell));
          for (let i = 1; i < spanOf(cell); i += 1) out.push('');
        }
        while (out.length < width) out.push('');
        return out;
      });
      const line = (cells: ReadonlyArray<string>): string => `| ${cells.join(' | ')} |`;
      const [head, ...body] = grid;
      const separator = line(Array.from({ length: width }, () => '---'));
      return `\n\n${[line(head ?? []), separator, ...body.map(line)].join('\n')}\n\n`;
    },
  });
};

const buildService = (options: { readonly gfm: boolean }): TurndownService => {
  const td = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced',
    bulletListMarker: '-',
    emDelimiter: '_',
    strongDelimiter: '**',
  });
  if (options.gfm) {
    td.use(gfm);
    addHeaderlessTableRule(td);
  }
  td.remove(['script', 'style']);
  return td;
};

const errorMessageOf = (e: unknown): string => (e instanceof Error ? e.message : String(e));

const decodeBasicEntities = (s: string): string =>
  s
    .replaceAll(/&nbsp;/gi, ' ')
    .replaceAll(/&amp;/gi, '&')
    .replaceAll(/&lt;/gi, '<')
    .replaceAll(/&gt;/gi, '>')
    .replaceAll(/&quot;/gi, '"')
    .replaceAll(/&#39;/gi, "'")
    .replaceAll(/&apos;/gi, "'");

const BLOCK_END_TAGS = new Set(['/p', '/div', '/h1', '/h2', '/h3', '/h4', '/h5', '/h6', '/li', '/tr']);

const tagInsertsNewline = (rawTag: string): boolean => {
  const lower = rawTag.toLowerCase();
  if (lower.startsWith('br')) {
    if (lower === 'br' || lower === 'br/' || lower.startsWith('br ') || lower.startsWith('br/')) return true;
  }
  return BLOCK_END_TAGS.has(lower);
};

/**
 * Walk the HTML byte-by-byte and emit either the character (when outside a
 * tag) or a newline (when closing a block tag / hitting <br>). Avoids the
 * regex `<[^>]*>` pattern that sonarjs flags as super-linear-backtracking
 * vulnerable, and skips <script>/<style> bodies in the same pass.
 */
const stripHtmlToText = (html: string): string => {
  let out = '';
  let cursor = 0;
  while (cursor < html.length) {
    const lt = html.indexOf('<', cursor);
    if (lt === -1) {
      out += html.slice(cursor);
      break;
    }
    out += html.slice(cursor, lt);
    const gt = html.indexOf('>', lt + 1);
    if (gt === -1) break;
    const rawTag = html.slice(lt + 1, gt);
    const lowerTag = rawTag.toLowerCase();
    if (rawTag.startsWith('!--')) {
      const end = html.indexOf('-->', lt + 4);
      cursor = end === -1 ? html.length : end + 3;
      continue;
    }
    if (lowerTag.startsWith('script') || lowerTag.startsWith('style')) {
      const tagName = lowerTag.startsWith('script') ? 'script' : 'style';
      const closer = html.toLowerCase().indexOf(`</${tagName}>`, gt + 1);
      cursor = closer === -1 ? html.length : closer + tagName.length + 3;
      continue;
    }
    if (tagInsertsNewline(rawTag)) out += '\n';
    cursor = gt + 1;
  }
  return decodeBasicEntities(out)
    .replaceAll(/[ \t]+/g, ' ')
    .replaceAll(/\n{3,}/g, '\n\n')
    .trim();
};

// Outlook and Word emit a header label and its colon as SEPARATE bold runs
// (`<b>发件人</b><b>:</b>`), and turndown faithfully closes and reopens each,
// so the `**发件人****:**` seam shows as stray asterisks wherever the renderer
// does not merge it. Two directly-adjacent SAME-tag runs are semantically one
// run, so the close+open boundary is dropped from the HTML before turndown
// ever sees it: turndown then emits a single `**发件人:**`. Working on the HTML
// source, not the markdown output, is deliberate — a blunt `replaceAll('****',
// '')` on the output also deletes a literal `****` inside a code span, a fenced
// block, or a URL (turndown does not escape those), and collapses nested bold
// (`**x**` -> ``) to nothing.
//
// Only zero-whitespace, same-tag adjacency matches (`\1` backreference): a real
// gap (`</b> <b>` = two bold words) is left alone, and a mixed `</b><strong>`
// pair is left alone rather than spliced into malformed markup.
const ADJACENT_BOLD_RUNS = /<\/(b|strong)><\1(?:\s[^>]*)?>/gi;
const mergeAdjacentBoldRuns = (html: string): string => html.replace(ADJACENT_BOLD_RUNS, '');

const htmlToMarkdown = (html: string): Result<string, GraphError> => {
  const merged = mergeAdjacentBoldRuns(html);
  try {
    return ok(buildService({ gfm: true }).turndown(merged));
  } catch (gfmError: unknown) {
    try {
      const md = buildService({ gfm: false }).turndown(merged);
      const note = `> _GFM table conversion failed: ${errorMessageOf(gfmError)}; tables flattened to paragraphs_`;
      return ok(`${note}\n\n${md}`);
    } catch (coreError: unknown) {
      const fallback = stripHtmlToText(html);
      const note = `> _markdown conversion failed: ${errorMessageOf(coreError)}; showing stripped HTML body_`;
      return ok(fallback.length > 0 ? `${note}\n\n${fallback}` : note);
    }
  }
};

export { htmlToMarkdown };
