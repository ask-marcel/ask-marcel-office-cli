declare module 'word-extractor' {
  // Minimal slice of the word-extractor surface the legacy-doc adapter uses:
  // `new WordExtractor().extract(buffer)` → a document whose `.getBody()` is the body text.
  type WordDocument = {
    getBody: () => string;
  };
  type WordExtractorInstance = {
    extract: (source: Buffer | string) => Promise<WordDocument>;
  };
  type WordExtractorConstructor = new () => WordExtractorInstance;
  const WordExtractor: WordExtractorConstructor;
  export default WordExtractor;
}
