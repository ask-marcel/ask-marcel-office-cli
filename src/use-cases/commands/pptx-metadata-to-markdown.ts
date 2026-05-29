import type { CustomProp, ExternalRel } from './ooxml-metadata.ts';
import { renderKv, renderTable } from './ooxml-metadata-to-markdown.ts';
import type { CommentAuthor, PptxComment, PptxMetadata, Slide, SlideTag } from './pptx-metadata.ts';

/**
 * Pure renderer: PptxMetadata → standalone `## PPTX metadata` document.
 * Uses the shared OOXML render primitives and assembles the
 * presentation-specific section order on top. Empty sections emit `_(none)_`.
 */

const renderCustom = (props: ReadonlyArray<CustomProp>): string =>
  renderTable(
    props.map((p) => [p.name, p.value]),
    ['name', 'value']
  );

const renderRels = (rels: ReadonlyArray<ExternalRel>): string =>
  renderTable(
    rels.map((r) => [r.source, r.type, r.target]),
    ['source', 'type', 'target']
  );

const renderTags = (tags: ReadonlyArray<SlideTag>): string =>
  renderTable(
    tags.map((t) => [t.source, t.name, t.value]),
    ['source', 'name', 'value']
  );

const renderAuthors = (authors: ReadonlyArray<CommentAuthor>): string =>
  renderTable(
    authors.map((a) => [a.id, a.name, a.initials]),
    ['id', 'name', 'initials']
  );

const renderComments = (comments: ReadonlyArray<PptxComment>): string =>
  renderTable(
    comments.map((c) => [c.author, c.date, c.text]),
    ['author', 'date', 'text']
  );

const renderSlides = (slides: ReadonlyArray<Slide>): string =>
  renderTable(
    slides.map((s) => [s.name, s.hidden ? 'hidden' : 'visible', s.title, s.notes]),
    ['slide', 'visibility', 'title', 'speaker notes']
  );

const formatPptxMetadata = (meta: PptxMetadata): string => {
  const sections: ReadonlyArray<readonly [string, string]> = [
    ['Core properties', renderKv(meta.core)],
    ['Application properties', renderKv(meta.app)],
    ['Custom properties', renderCustom(meta.custom)],
    ['External relationships', renderRels(meta.externalRels)],
    ['Slide tags', renderTags(meta.slideTags)],
    ['Comment authors', renderAuthors(meta.commentAuthors)],
    ['Comments', renderComments(meta.comments)],
    ['Slides', renderSlides(meta.slides)],
  ];
  const body = sections.map(([title, content]) => `### ${title}\n\n${content}`).join('\n\n');
  return `## PPTX metadata\n\n${body}\n`;
};

export { formatPptxMetadata };
