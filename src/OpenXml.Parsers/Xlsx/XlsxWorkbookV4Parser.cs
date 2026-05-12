using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXml.Parsers.Xlsx;

public static class XlsxWorkbookV4Parser
{
    public static XlsxWorkbookParseResultV4 Parse(string filePath)
    {
        var v3 = XlsxWorkbookV3Parser.Parse(filePath);
        var warnings = v3.Warnings.ToList();
        var sheetLayouts = ParseSheetLayouts(filePath, warnings);

        var sheets = v3.Workbook.Sheets
            .Select(sheet =>
            {
                sheetLayouts.TryGetValue(sheet.Id, out var layout);
                return new XlsxSheetV4(
                    sheet.Id,
                    sheet.Name,
                    sheet.Index,
                    sheet.Dimension,
                    sheet.Cells,
                    layout?.MergedCells ?? [],
                    layout?.Rows ?? [],
                    layout?.Columns ?? []);
            })
            .ToList();

        return new XlsxWorkbookParseResultV4(
            XlsxWorkbookParseResultV4.CurrentSchemaVersion,
            v3.Source,
            new XlsxWorkbookV4(sheets, v3.Workbook.Styles),
            warnings);
    }

    private static Dictionary<string, SheetLayout> ParseSheetLayouts(string filePath, List<string> warnings)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("Workbook part is missing.");

        var layouts = new Dictionary<string, SheetLayout>(StringComparer.OrdinalIgnoreCase);
        var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];

        foreach (var sheet in sheets)
        {
            var relationshipId = sheet.Id?.Value;
            var sheetName = sheet.Name?.Value ?? relationshipId ?? "Unknown";
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            if (workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var worksheet = worksheetPart.Worksheet;
            if (worksheet is null)
            {
                warnings.Add($"Sheet '{sheetName}' has no worksheet payload for layout metadata.");
                continue;
            }

            layouts[relationshipId] = new SheetLayout(
                ParseMergedCells(worksheet, sheetName, warnings),
                ParseRows(worksheet),
                ParseColumns(worksheet));
        }

        return layouts;
    }

    private static List<XlsxMergedCellV4> ParseMergedCells(Worksheet worksheet, string sheetName, List<string> warnings)
    {
        var mergedCells = new List<XlsxMergedCellV4>();

        foreach (var mergeCell in worksheet.Elements<MergeCells>().SelectMany(container => container.Elements<MergeCell>()))
        {
            var reference = mergeCell.Reference?.Value;
            if (string.IsNullOrWhiteSpace(reference))
            {
                warnings.Add($"Sheet '{sheetName}' contains a merged cell without a reference.");
                continue;
            }

            var addresses = reference.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var startAddress = addresses[0];
            var endAddress = addresses.Length > 1 ? addresses[1] : addresses[0];

            if (!TryParseAddress(startAddress, out var startRow, out var startColumn) ||
                !TryParseAddress(endAddress, out var endRow, out var endColumn))
            {
                warnings.Add($"Sheet '{sheetName}' contains an invalid merged cell reference: '{reference}'.");
                continue;
            }

            mergedCells.Add(new XlsxMergedCellV4(
                reference,
                startAddress,
                endAddress,
                startRow,
                startColumn,
                endRow,
                endColumn));
        }

        return mergedCells;
    }

    private static List<XlsxRowV4> ParseRows(Worksheet worksheet)
    {
        return worksheet.Descendants<Row>()
            .Where(row => row.RowIndex?.Value is not null)
            .Select(row => new XlsxRowV4(
                row.RowIndex!.Value,
                row.Height?.Value,
                row.CustomHeight?.Value == true,
                row.Hidden?.Value == true,
                row.StyleIndex?.Value,
                row.OutlineLevel?.Value,
                row.Collapsed?.Value == true))
            .ToList();
    }

    private static List<XlsxColumnV4> ParseColumns(Worksheet worksheet)
    {
        return worksheet.Elements<Columns>()
            .SelectMany(columns => columns.Elements<Column>())
            .Select(column => new XlsxColumnV4(
                column.Min?.Value ?? 0,
                column.Max?.Value ?? 0,
                column.Width?.Value,
                column.CustomWidth?.Value == true,
                column.BestFit?.Value == true,
                column.Hidden?.Value == true,
                column.Style?.Value,
                column.OutlineLevel?.Value,
                column.Collapsed?.Value == true))
            .ToList();
    }

    private static bool TryParseAddress(string address, out uint row, out int column)
    {
        row = 0;
        column = 0;

        var letters = new string(address.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(address.SkipWhile(char.IsLetter).TakeWhile(char.IsDigit).ToArray());
        if (letters.Length == 0 || digits.Length == 0)
        {
            return false;
        }

        foreach (var letter in letters.ToUpperInvariant())
        {
            column *= 26;
            column += letter - 'A' + 1;
        }

        return uint.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out row);
    }

    private sealed record SheetLayout(
        IReadOnlyList<XlsxMergedCellV4> MergedCells,
        IReadOnlyList<XlsxRowV4> Rows,
        IReadOnlyList<XlsxColumnV4> Columns);
}
