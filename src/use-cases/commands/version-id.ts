/**
 * Graph's historical-version endpoints template `{version-id}` as a string
 * that must look like a stringified float (`"79.0"`, `"4.0"`). An LLM
 * (reasonably) types `79`, and the response is wildly inconsistent across
 * the three sibling commands:
 *
 *   - download-drive-item-version-content        → `invalidRequest`
 *   - download-drive-item-version-as-markdown    → `invalidRequest`
 *   - download-drive-item-version-as-pdf         → accepts (Graph quirk)
 *
 * Audit round-6 §1.4 caught the inconsistency. Normalize at the CLI
 * boundary: a pure-integer version-id gets `.0` appended before path
 * substitution, so all three commands accept both spellings.
 *
 * Already-floated values (`79.0`, `12.5`, `4.1`) and non-numeric values
 * (a bad input the schema will then reject) pass through unchanged.
 */
export const normalizeVersionId = (raw: string): string => (/^\d+$/.test(raw) ? `${raw}.0` : raw);
