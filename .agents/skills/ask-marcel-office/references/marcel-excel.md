# Excel commands

## get-excel-chart-image
Render a chart on an Excel worksheet as a PNG (base64). Calls Graph's chart `Image()` function (natural size, aspect-preserving) so a vision-capable LLM can read the plotted data itself — not just the chart's title / position metadata that `list-excel-worksheet-charts` returns. The chart id or name 
Required: --drive-id --item-id --worksheet-id --chart-id
Example: ask-marcel get-excel-chart-image --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1' --chart-id 'Chart 1' --output-path ./chart.png
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/charts/{chart-id}/Image(width=0,height=0,fittingMode='Fit')

## get-excel-range
Get the cell values, formulas, and formats of a specific Excel range (e.g. `A1:C10`). The CLI caps the in-flight range at 100 000 cells to prevent runaway responses — split absurd ranges (`ZZ999999:AAA1` etc.) into smaller bands.
Required: --drive-id --item-id --worksheet-id --address
Example: ask-marcel get-excel-range --drive-id 'b!1234' --item-id '01XLSX' --worksheet-id 'Sheet1' --address 'A1:C10'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/range(address='{address}')

## get-excel-table
Get the metadata (style, header row, total row) of a single named Excel table.
Required: --drive-id --item-id --table-id
Example: ask-marcel get-excel-table --drive-id 'b!1234' --item-id '01XLSX' --table-id 'Table1'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}

## get-excel-used-range
Return the worksheet's used range — the bounding box of every non-empty cell — as a single Excel range. The CLI ships a slim default that strips the redundant `text` / `numberFormat` / `formulas` 2D arrays Graph returns (mostly `"General"` repeated cell-by-cell), keeping `address` / `rowCount` / `co
Required: --drive-id --item-id --worksheet-id
Optional: --full --max-cells
Example: ask-marcel get-excel-used-range --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/usedRange()

## list-excel-comments
List the modern threaded comments anchored to cells in an Excel workbook (the New Comments feature, distinct from legacy notes). Each `workbookComment` has `content`, `contentType`, `task` state, plus replies via the comment's `replies` navigation.
Required: --drive-id --item-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-excel-comments --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/comments

## list-excel-defined-names
List the workbook's defined names (named ranges, named formulas, named constants). Each `workbookNamedItem` has `name`, `value` (the formula or address), `comment`, and `scope` (workbook or worksheet). Useful for understanding workbook structure before reading ranges.
Required: --drive-id --item-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-excel-defined-names --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/names

## list-excel-table-rows
List the data rows of a named Excel table (excluding the header row). Note: Graph silently ignores `$filter` and `$orderby` on this endpoint, so the CLI does not expose those flags — slice / sort client-side.
Required: --drive-id --item-id --table-id
Optional: --top --skip --select --expand
Example: ask-marcel list-excel-table-rows --drive-id 'b!1234' --item-id '01XLSX' --table-id 'Table1'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}/rows

## list-excel-tables
List the named tables across every worksheet in an Excel workbook. Note: Graph silently ignores `$filter` and `$orderby` on this endpoint, so the CLI does not expose those flags — slice / sort client-side.
Required: --drive-id --item-id
Optional: --top --skip --select --expand
Example: ask-marcel list-excel-tables --drive-id 'b!1234' --item-id '01XLSX'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/tables

## list-excel-worksheet-charts
List the charts on a worksheet. Each `workbookChart` has `id`, `name`, `height`, `width`, `top`, `left`. Use the chart's image endpoint (`.../charts/{id}/image()`) to render the chart as a base64 PNG.
Required: --drive-id --item-id --worksheet-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-excel-worksheet-charts --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/charts

## list-excel-worksheet-pivot-tables
List the pivot tables on a worksheet. Each `workbookPivotTable` has `name` and a navigation to its source `workbookWorksheet`. Useful for understanding analytical structure inside a workbook.
Required: --drive-id --item-id --worksheet-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-excel-worksheet-pivot-tables --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/pivotTables

## list-excel-worksheets
List the worksheets (tabs) inside an Excel workbook stored in OneDrive / SharePoint. Returns a clear "not an accessible Excel workbook" error if the item is a folder, non-.xlsx file, or sensitivity-label-blocked. Note: Graph silently ignores `$top`, `$filter`, and `$orderby` on this endpoint, so the
Required: --drive-id --item-id
Optional: --skip --select --expand
Example: ask-marcel list-excel-worksheets --drive-id 'b!1234' --item-id '01XLSX'
Graph: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets
