import { describe, expect, it } from 'bun:test';
import { extractSignatureBlock } from './signature-extractor.ts';

const QUOTE_TAIL = '<div id="appendonsend"></div><div id="divRplyFwdMsg"><b>From:</b> Robin Chen<br><b>Sent:</b> Monday</div>';

describe('lifting the sender’s own signature block out of a message they sent', () => {
  it('returns the whole signature with its nested layout divs intact', () => {
    const signature = '<div id="Signature"><div class="card"><div class="row"><b>Robin Chen</b></div><div class="row">Fabrikam</div></div></div>';
    const html = `<html><body><div>the message text</div>${signature}${QUOTE_TAIL}<div>the quoted original</div></body></html>`;

    // Three levels of nesting: a depth counter that stops one close tag early
    // would return a truncated card, and one that runs on would swallow the quote.
    expect(extractSignatureBlock(html)).toBe(signature);
  });

  it('ignores a signature that only appears inside the quoted history, so a colleague’s block is never returned as the sender’s', () => {
    const html = `<html><body><div>my reply</div>${QUOTE_TAIL}<div id="Signature">Alex Kim, Contoso</div></body></html>`;

    expect(extractSignatureBlock(html)).toBeUndefined();
  });

  it('returns the sender’s own signature and not the one quoted below it when a thread carries both', () => {
    const mine = '<div id="Signature">Robin Chen, Fabrikam</div>';
    const html = `<html><body><div>my reply</div>${mine}${QUOTE_TAIL}<div id="Signature">Alex Kim, Contoso</div></body></html>`;

    expect(extractSignatureBlock(html)).toBe(mine);
  });

  it('finds the block whatever the attribute order and tag casing Outlook happens to emit', () => {
    const signature = '<DIV style="font-family:Aptos" ID="Signature" dir="ltr">Robin Chen</DIV>';
    const html = `<html><body><div>text</div>${signature}</body></html>`;

    expect(extractSignatureBlock(html)).toBe(signature);
  });

  it('returns nothing for a message with no signature block at all, rather than guessing at one', () => {
    expect(extractSignatureBlock(`<html><body><div>just a message</div>${QUOTE_TAIL}</body></html>`)).toBeUndefined();
  });

  it('returns nothing when the signature markup is unbalanced, rather than running to the end of the document', () => {
    // The close tag is missing, so there is no honest end to the block; taking
    // the rest of the document would drag the quoted thread into the signature.
    const html = '<html><body><div id="Signature"><div class="card">Robin Chen</div></body></html>';

    expect(extractSignatureBlock(html)).toBeUndefined();
  });

  it('finds a signature in a message that quotes nothing, where the whole body is the sender’s own text', () => {
    const signature = '<div id="Signature">Robin Chen, Fabrikam</div>';

    expect(extractSignatureBlock(`<html><body><div>a fresh message</div>${signature}</body></html>`)).toBe(signature);
  });
});
