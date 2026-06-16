import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

type SheetCsv = { readonly name: string; readonly csv: string };

const readSheetsAsCsv = async (bytes: Uint8Array): Promise<Result<ReadonlyArray<SheetCsv>, GraphError>> => {
  try {
    // Lazy-load SheetJS (~181 ms init — the single heaviest dep) so non-Excel
    // commands don't pay for it at startup; loads on the first .xlsx/.xls/.csv
    // conversion only (QA-012 cold-start).
    const XLSX = await import('xlsx');
    const workbook = XLSX.read(bytes, { type: 'array' });
    const sheets = workbook.SheetNames.map((name) => {
      const sheet = workbook.Sheets[name];
      if (!sheet) return { name, csv: '' };
      // `blankrows: false` drops fully-blank rows — Excel routinely pads the used
      // range far past the real data, and those empty rows otherwise emit millions
      // of bare-`,` lines (a 49 MB workbook crashed the markdown builder this way).
      return { name, csv: XLSX.utils.sheet_to_csv(sheet, { blankrows: false }) };
    });
    return ok(sheets);
  } catch (e) {
    return err({ type: 'api_error', status: 500, message: `xlsx parse failed: ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { readSheetsAsCsv };
export type { SheetCsv };
