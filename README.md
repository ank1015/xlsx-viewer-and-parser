# OpenXML Parser Workspace

This workspace is set up for building parsers around Microsoft Office Open XML files such as `.xlsx`, `.docx`, and `.pptx`.

## Requirements

- .NET SDK 10.0.107 installed through Homebrew
- OpenXML SDK package: `DocumentFormat.OpenXml` 3.5.1

If another tool needs the .NET root explicitly, use:

```zsh
export DOTNET_ROOT="/opt/homebrew/opt/dotnet/libexec"
```

## Projects

- `src/OpenXml.Parsers`: reusable parsing library
- `src/OpenXml.Cli`: command-line runner for inspecting files and emitting parser JSON
- `tests/OpenXml.Parsers.Tests`: smoke tests that create real `.docx` and `.xlsx` files and inspect them

## Common Commands

```zsh
dotnet build OpenXmlWorkspace.slnx
dotnet test OpenXmlWorkspace.slnx
dotnet run --project src/OpenXml.Cli -- xlsx /path/to/file.xlsx
dotnet run --project src/OpenXml.Cli -- xlsx-assets /path/to/file.xlsx /path/to/output-dir
dotnet run --project src/OpenXml.Cli -- xlsx-v5 /path/to/file.xlsx
dotnet run --project src/OpenXml.Cli -- xlsx-v4 /path/to/file.xlsx
dotnet run --project src/OpenXml.Cli -- xlsx-v3 /path/to/file.xlsx
dotnet run --project src/OpenXml.Cli -- xlsx-v2 /path/to/file.xlsx
dotnet run --project src/OpenXml.Cli -- xlsx-v1 /path/to/file.xlsx
dotnet run --project src/OpenXml.Cli -- inspect /path/to/file.docx
```

## CLI Install

The packaged CLI is named `heysnap-xlsxl`.

Install the latest Linux release:

```sh
curl -fsSL https://github.com/ank1015/xlsx-viewer-and-parser/releases/latest/download/install.sh | sh
```

Install a specific version:

```sh
curl -fsSL https://github.com/ank1015/xlsx-viewer-and-parser/releases/download/v0.1.0/install.sh | HEYSNAP_XLSXL_VERSION=v0.1.0 sh
```

The installer detects `linux-x64` and `linux-arm64`, downloads the matching self-contained binary, and installs it to `/usr/local/bin` by default. Override the install location with `HEYSNAP_XLSXL_INSTALL_DIR`.

Build release artifacts locally:

```zsh
scripts/package-cli.sh 0.1.0
```

This creates release assets under `artifacts/release/v0.1.0`.

## Demo Server

Run the local demo server on port 8000:

```zsh
node server.js
```

Call the spreadsheet parser by demo file number:

```zsh
curl -X POST http://localhost:8000/xlsx/1
```

You can also send the number in the body:

```zsh
curl -X POST http://localhost:8000/xlsx \
  -H 'Content-Type: application/json' \
  -d '{"number":1}'
```

`GET /` returns the available number-to-file mapping.

## Demo Client

Run the Vite client on port 3000:

```zsh
cd client
npm install
npm run dev
```

Open `http://localhost:3000/1`, `http://localhost:3000/2`, and so on. Each page calls the matching demo server endpoint when you click `Fetch`, then renders the workbook in a data grid. Sheet buttons appear above the grid, and selecting a formula cell shows its formula text in the formula bar.

## Current Parser Surface

`XlsxWorkbookV6Parser.Parse(path)` emits `xlsx-workbook/v6` JSON-ready data:

- source file metadata
- workbook sheets
- sheet id, name, index, and dimension
- merged cell ranges with start/end addresses and numeric row/column bounds
- explicit row metadata including height, hidden state, style index, outline level, and collapsed state
- explicit column range metadata including width, hidden state, style index, best-fit, outline level, and collapsed state
- table definitions with relationship id, id, name, display name, range, bounds, row-count metadata, table columns, auto-filter metadata, and style flags
- worksheet drawing metadata
- chart metadata including type, title, anchor, series ranges, and cached series values
- image metadata including content type, file name, description, and anchor
- optional relative asset paths for extracted chart XML and image files
- sparse cells keyed by address
- cell address, row, column, type, raw value, parsed value, display value, and style index
- basic cell types: string, number, date, boolean, blank, and error
- formula text and metadata for formula cells
- resolved shared formula text for dependent shared-formula cells
- formula cached result raw value, parsed value, display value, and value type
- workbook style tables for number formats, fonts, fills, borders, and cell formats
- basic formatted display values for styled numbers and dates
- parse warnings

`XlsxWorkbookV6Parser.Parse(path, assetOutputDirectory)` extracts chart XML into `assets/charts` and images into `assets/images`, using relative asset paths in `workbook.json`.

`XlsxWorkbookV5Parser.Parse(path)` remains available for table definitions without drawing/chart/image metadata.

`XlsxWorkbookV4Parser.Parse(path)` remains available for sheet layout metadata without table definitions.

`XlsxWorkbookV3Parser.Parse(path)` remains available for style-aware parsing without sheet layout metadata.

`XlsxWorkbookV2Parser.Parse(path)` remains available for formula-aware parsing without style tables.

`XlsxWorkbookV1Parser.Parse(path)` remains available for the initial cells-only schema.

`OpenXmlInspector.Inspect(path)` currently supports:

- `.docx`: paragraph count and text preview
- `.xlsx`: sheet count, sheet names, and first rows preview
- `.pptx`: package open check and slide count

The inspector API returns an `OpenXmlInspectionResult`, which is intentionally simple for quick summaries. UI-facing parser work should use the versioned parser models.
