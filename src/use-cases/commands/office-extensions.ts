/**
 * Single source of truth for the OOXML extension families the markdown
 * dispatchers route together. Each family shares one parser/extractor: the
 * macro-enabled (`*m`) and template (`*tx` / `*tm`) variants are structurally
 * identical to their base format, so they alias onto the same path. Keeping
 * the lists here (not duplicated per dispatcher) means a new variant is added
 * in exactly one place.
 */

const DOCX_FAMILY: ReadonlySet<string> = new Set(['docx', 'docm', 'dotx', 'dotm']);
const XLSX_FAMILY: ReadonlySet<string> = new Set(['xlsx', 'xlsm', 'xltx', 'xltm']);
const PPTX_FAMILY: ReadonlySet<string> = new Set(['pptx', 'pptm', 'potx', 'potm']);

export { DOCX_FAMILY, PPTX_FAMILY, XLSX_FAMILY };
