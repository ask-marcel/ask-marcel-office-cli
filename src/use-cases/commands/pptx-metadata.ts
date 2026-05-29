import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractAppProps, extractCoreProps, extractCustomProps, extractExternalRels } from './ooxml-metadata.ts';
import type { CustomProp, ExternalRel } from './ooxml-metadata.ts';
import { attrOf, findAll, parseXml } from './ooxml-xml-walker.ts';
import { extractCommentAuthors, extractComments } from './pptx-comments.ts';
import type { CommentAuthor, PptxComment } from './pptx-comments.ts';
import { extractSlides } from './pptx-slides.ts';
import type { Slide } from './pptx-slides.ts';

/**
 * Pulls the side-channel content out of a .pptx zip — everything a user can
 * author that the rendered slide PDF never shows: core / app / custom doc
 * properties, external relationships, slide tags (key/value), comment authors
 * and comments (legacy + modern), and per-slide title / speaker notes / hidden
 * flag. Package-level parts come from the shared ooxml-metadata module; this
 * file owns the presentation-specific parts.
 */

type PropMap = Readonly<Record<string, string>>;
type SlideTag = { readonly source: string; readonly name: string; readonly value: string };

type PptxMetadata = {
  readonly core: PropMap;
  readonly app: PropMap;
  readonly custom: ReadonlyArray<CustomProp>;
  readonly externalRels: ReadonlyArray<ExternalRel>;
  readonly slideTags: ReadonlyArray<SlideTag>;
  readonly commentAuthors: ReadonlyArray<CommentAuthor>;
  readonly comments: ReadonlyArray<PptxComment>;
  readonly slides: ReadonlyArray<Slide>;
};

const tagsInPart = (root: unknown, source: string): ReadonlyArray<SlideTag> =>
  findAll(root, 'p:tag').map((tag) => ({ source, name: attrOf(tag, 'name'), value: attrOf(tag, 'val') }));

const extractSlideTags = (zip: OoxmlZip): ReadonlyArray<SlideTag> => {
  const out: Array<SlideTag> = [];
  for (const path of zip.list().filter((p) => /^ppt\/tags\/tag\d+\.xml$/.test(p))) out.push(...tagsInPart(parseXml(zip.read(path)), path));
  return out;
};

const extractPptxMetadata = async (bytes: Uint8Array): Promise<Result<PptxMetadata, GraphError>> => {
  const zipR = await openOoxmlZip(bytes);
  if (!zipR.ok) return zipR;
  const zip = zipR.value;
  const commentAuthors = extractCommentAuthors(zip);
  return ok({
    core: extractCoreProps(zip),
    app: extractAppProps(zip),
    custom: extractCustomProps(zip),
    externalRels: extractExternalRels(zip),
    slideTags: extractSlideTags(zip),
    commentAuthors,
    comments: extractComments(zip, commentAuthors),
    slides: extractSlides(zip),
  });
};

export { extractPptxMetadata };
export type { CommentAuthor, PptxComment, PptxMetadata, Slide, SlideTag };
