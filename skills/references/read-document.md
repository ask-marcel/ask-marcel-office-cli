# Read a document in full

First get the file's `drive-id` + `item-id`.

**From a search hit** — the hit carries several ids; use exactly the two marked (picking the wrong one 404s):

```
value:
  id: 01ABC…                 ← --item-id   (the id at the SAME level as `name`)
  name: Q3 Budget.xlsx
  parentReference:
    driveId: b!xY…z          ← --drive-id
    id: 01GHI…               ✗ parent FOLDER — not the file
  listItem:
    id: 9f3c…                ✗ list-item id — not the file
```

**From a sharing URL** (`*.sharepoint.com` or `1drv.ms` "Copy link"):

```bash
ask-marcel-office resolve-drive-share-link --url '<sharing url>'
```

Returns `driveId` + `itemId` + `tenantId`. If `tenantId` isn't yours (a share from another org), add `--tenant-id '<tenantId>'` to every download/convert command for that file.

**Then read it, by type:**

- **PDF / CSV / plain text** — already readable, download raw: `download-drive-item-content --drive-id … --item-id … --output-path file.<real ext>`
- **Word / Excel / OpenDocument** — `download-drive-item-as-markdown --drive-id … --item-id … --include-metadata true`. `--include-metadata true` surfaces comments, tracked changes, hidden text. Leave images as their default `[image: <alt>]` placeholders — never pass `--inline-images true`; base64 in markdown is bytes you cannot see. When a placeholder looks content-bearing (screenshot, diagram), `extract-drive-item-images --drive-id … --item-id … --output-dir ./imgs` writes the full-resolution originals as files — Read them; an embedded image (a cost table saved as a picture, a diagram) holds real content (an Excel chart renders via `get-excel-chart-image`). A scrambled **Word/OpenDocument** conversion (scanned pages, layout turned to soup) is not worth fighting — fall back to `download-drive-item-as-pdf` and read the PDF. A messy Excel is different: go sheet-by-sheet (next bullet), never to PDF (pages slice the sheets).
- **Big or many-sheeted Excel** — go sheet by sheet: `list-excel-worksheets` then `get-excel-used-range --worksheet-id '<name>'`. Pass `--full true` to get formulas and value types alongside values — it shows directly whether a total is computed or hand-typed. Named tables read via `list-excel-tables` → `list-excel-table-rows`. Also run `download-drive-item-as-markdown --include-metadata true` once for the `## Workbook metadata` block — cell comments (often the "why" behind a number), hidden sheets, defined names — which the sheet reads don't include.
- **PowerPoint / anything where layout matters** — `download-drive-item-as-pdf --drive-id … --item-id … --output-path deck.pdf`, then read the PDF.
- **Zip archives** — one call unzips and converts every file inside (legacy GBK/CP437 entry names decoded): `convert-drive-item-zip-to-markdown` (in OneDrive/SharePoint), `convert-mail-attachment-zip-to-markdown` (mail), `convert-local-file-to-markdown --path ./archive.zip` (disk). It returns the text content and lists images and scanned/image-only PDFs without unpacking them. To read those scanned entries (a stamped invoice, a bank passbook, a registration cert): `download-drive-item-content --drive-id … --item-id … --output-path archive.zip`, unzip locally, then Read the image and PDF files — Read renders PDF pages visually, scans included. Triage from the converter's scan-only list: open the files the question needs and the pages that carry the data; skip boilerplate and duplicate copies.
- **Follow references out of the doc:** `extract-sharepoint-links-in-documents --drive-id … --item-id …`
- **A file you saved to disk during this task** (an attachment or download; works logged-out): `convert-local-file-to-markdown --path './report.docx'`

**When the numbers matter** — these hold for any sheet, whatever route produced the markdown:

- **Formula errors survive conversion.** `#REF!`, `#N/A`, `#VALUE!`, `#DIV/0!` come through as-is; when a summary cell shows one, recompute the figure from the detail rows and say that you did.
- **Reconcile hand-typed totals.** Check a grand total against its rows and flag any mismatch; `--full true` on `get-excel-used-range` shows whether a cell is a formula or a typed value.
- **Count with a script, never by eye.** How many items carry each status, how many are still open: `grep -c` / `awk` over the converted markdown. Hand-counting a few hundred rows in a wide sheet is unreliable and two reads rarely agree; a one-line filter is exact and repeatable.
