using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXml.Parsers.Xlsx;

public static class XlsxWorkbookV5Parser
{
    public static XlsxWorkbookParseResultV5 Parse(string filePath)
    {
        var v4 = XlsxWorkbookV4Parser.Parse(filePath);
        var warnings = v4.Warnings.ToList();
        var sheetTables = ParseSheetTables(filePath, warnings);

        var sheets = v4.Workbook.Sheets
            .Select(sheet =>
            {
                sheetTables.TryGetValue(sheet.Id, out var tables);
                return new XlsxSheetV5(
                    sheet.Id,
                    sheet.Name,
                    sheet.Index,
                    sheet.Dimension,
                    sheet.Cells,
                    sheet.MergedCells,
                    sheet.Rows,
                    sheet.Columns,
                    tables ?? []);
            })
            .ToList();

        return new XlsxWorkbookParseResultV5(
            XlsxWorkbookParseResultV5.CurrentSchemaVersion,
            v4.Source,
            new XlsxWorkbookV5(sheets, v4.Workbook.Styles),
            warnings);
    }

    private static Dictionary<string, IReadOnlyList<XlsxTableV5>> ParseSheetTables(string filePath, List<string> warnings)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("Workbook part is missing.");

        var tablesBySheetId = new Dictionary<string, IReadOnlyList<XlsxTableV5>>(StringComparer.OrdinalIgnoreCase);
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

            var tables = new List<XlsxTableV5>();
            foreach (var tablePart in worksheetPart.TableDefinitionParts)
            {
                var tableRelationshipId = worksheetPart.GetIdOfPart(tablePart);
                var table = tablePart.Table;
                if (table is null)
                {
                    warnings.Add($"Sheet '{sheetName}' has table relationship '{tableRelationshipId}' with no table payload.");
                    continue;
                }

                tables.Add(ParseTable(table, tableRelationshipId, sheetName, warnings));
            }

            tablesBySheetId[relationshipId] = tables;
        }

        return tablesBySheetId;
    }

    private static XlsxTableV5 ParseTable(
        Table table,
        string relationshipId,
        string sheetName,
        List<string> warnings)
    {
        var reference = table.Reference?.Value;
        var startAddress = default(string);
        var endAddress = default(string);
        uint? startRow = null;
        int? startColumn = null;
        uint? endRow = null;
        int? endColumn = null;

        if (!string.IsNullOrWhiteSpace(reference))
        {
            var addresses = reference.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            startAddress = addresses[0];
            endAddress = addresses.Length > 1 ? addresses[1] : addresses[0];

            if (TryParseAddress(startAddress, out var parsedStartRow, out var parsedStartColumn) &&
                TryParseAddress(endAddress, out var parsedEndRow, out var parsedEndColumn))
            {
                startRow = parsedStartRow;
                startColumn = parsedStartColumn;
                endRow = parsedEndRow;
                endColumn = parsedEndColumn;
            }
            else
            {
                warnings.Add($"Sheet '{sheetName}' has table '{table.DisplayName?.Value ?? relationshipId}' with invalid reference '{reference}'.");
            }
        }

        return new XlsxTableV5(
            relationshipId,
            table.Id?.Value,
            table.Name?.Value,
            table.DisplayName?.Value,
            reference,
            startAddress,
            endAddress,
            startRow,
            startColumn,
            endRow,
            endColumn,
            table.HeaderRowCount?.Value ?? 1,
            table.TotalsRowCount?.Value ?? 0,
            table.InsertRow?.Value == true,
            table.Published?.Value == true,
            ParseTableColumns(table),
            ParseAutoFilter(table.AutoFilter),
            ParseTableStyle(table.TableStyleInfo));
    }

    private static List<XlsxTableColumnV5> ParseTableColumns(Table table)
    {
        return table.TableColumns?.Elements<TableColumn>()
            .Select(column => new XlsxTableColumnV5(
                column.Id?.Value ?? 0,
                column.Name?.Value ?? string.Empty,
                column.TotalsRowLabel?.Value,
                column.TotalsRowFunction?.InnerText,
                column.CalculatedColumnFormula?.Text,
                column.TotalsRowFormula?.Text))
            .ToList() ?? [];
    }

    private static XlsxAutoFilterV5? ParseAutoFilter(AutoFilter? autoFilter)
    {
        if (autoFilter is null)
        {
            return null;
        }

        var filterColumns = autoFilter.Elements<FilterColumn>()
            .Select(column => new XlsxFilterColumnV5(
                column.ColumnId?.Value ?? 0,
                column.HiddenButton?.Value == true,
                column.ShowButton?.Value != false,
                column.ChildElements.Select(element => element.LocalName).ToList()))
            .ToList();

        return new XlsxAutoFilterV5(autoFilter.Reference?.Value, filterColumns);
    }

    private static XlsxTableStyleV5? ParseTableStyle(TableStyleInfo? tableStyleInfo)
    {
        if (tableStyleInfo is null)
        {
            return null;
        }

        return new XlsxTableStyleV5(
            tableStyleInfo.Name?.Value,
            tableStyleInfo.ShowFirstColumn?.Value == true,
            tableStyleInfo.ShowLastColumn?.Value == true,
            tableStyleInfo.ShowRowStripes?.Value == true,
            tableStyleInfo.ShowColumnStripes?.Value == true);
    }

    private static bool TryParseAddress(string address, out uint row, out int column)
    {
        row = 0;
        column = 0;

        var normalized = address.Replace("$", string.Empty, StringComparison.Ordinal);
        var letters = new string(normalized.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(normalized.SkipWhile(char.IsLetter).TakeWhile(char.IsDigit).ToArray());
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
}
