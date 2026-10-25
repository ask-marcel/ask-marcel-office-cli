import { describe, expect, it } from 'bun:test';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { buildAdversarialPptx, buildMinimalPptx, buildRichPptx } from '../../test-helpers/office-fixtures.ts';
import { extractSlides } from './pptx-slides.ts';
import type { Slide } from './pptx-slides.ts';

const slidesOf = async (bytes: Uint8Array): Promise<ReadonlyArray<Slide>> => {
  const zip = await openOoxmlZip(bytes);
  if (!zip.ok) throw new Error('failed to open fixture zip');
  return extractSlides(zip.value);
};

describe('extractSlides', () => {
  it('extracts each slide’s name, hidden flag, title, body text, and resolved speaker notes — in slide-number order', async () => {
    expect(await slidesOf(await buildRichPptx())).toEqual([
      {
        name: 'slide1.xml',
        hidden: false,
        title: 'Quarterly Review',
        notes: 'Remember to mention the Q3 shortfall and the ACME contract renewal.',
        text: 'Quarterly Review\nsee the portal',
      },
      {
        name: 'slide2.xml',
        hidden: true, // <p:sld show="0">
        title: 'Internal Only — Do Not Present', // resolved from a ctrTitle placeholder, not just `title`
        notes: '', // no notesSlide relationship
        text: 'Internal Only — Do Not Present',
      },
    ]);
  });

  it('returns empty title and notes for a barebones untitled slide with no notesSlide, keeping its body text', async () => {
    expect(await slidesOf(await buildMinimalPptx())).toEqual([{ name: 'slide1.xml', hidden: false, title: '', notes: '', text: 'plain content' }]);
  });

  it('sorts by slide number (not list/lexical order), skips non-title placeholders, filters blank paragraphs, and resolves a non-first notesSlide rel', async () => {
    // slide10 is added before slide2 in the zip and is two-digit, so a broken sort would surface it first.
    expect(await slidesOf(await buildAdversarialPptx())).toEqual([
      { name: 'slide2.xml', hidden: true, title: 'Second', notes: 'note one note two', text: 'Second' },
      { name: 'slide10.xml', hidden: false, title: 'The Title', notes: '', text: 'body first\nThe Title' },
    ]);
  });
});
