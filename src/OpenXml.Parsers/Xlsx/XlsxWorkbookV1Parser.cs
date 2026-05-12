using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXml.Parsers.Xlsx;

public static class XlsxWorkbookV1Parser
{
    public static XlsxWorkbookParseResultV1 Parse(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("An XLSX file path is required.", nameof(filePath));
        }

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("XLSX file was not found.", filePath);
        }

        var warnings = new List<string>();
        var fileInfo = new FileInfo(filePath);

        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("Workbook part is missing.");

        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
        var parsedSheets = new List<XlsxSheetV1>();

        for (var sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++)
        {
            var sheet = sheets[sheetIndex];
            var relationshipId = sheet.Id?.Value;
            var sheetName = sheet.Name?.Value ?? $"Sheet{sheetIndex + 1}";

            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                warnings.Add($"Sheet '{sheetName}' does not have a relationship id and was skipped.");
                continue;
            }

            if (workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                warnings.Add($"Sheet '{sheetName}' points to a non-worksheet part and was skipped.");
                continue;
            }

            parsedSheets.Add(ParseSheet(
                worksheetPart,
                relationshipId,
                sheetName,
                sheetIndex,
                sharedStrings,
                warnings));
        }

        return new XlsxWorkbookParseResultV1(
            XlsxWorkbookParseResultV1.CurrentSchemaVersion,
            new XlsxSourceInfoV1(fileInfo.Name, fileInfo.FullName, fileInfo.Length),
            new XlsxWorkbookV1(parsedSheets),
            warnings);
    }

    private static XlsxSheetV1 ParseSheet(
        WorksheetPart worksheetPart,
        string relationshipId,
        string sheetName,
        int sheetIndex,
        SharedStringTable? sharedStrings,
        List<string> warnings)
    {
        var worksheet = worksheetPart.Worksheet;
        if (worksheet is null)
        {
            warnings.Add($"Sheet '{sheetName}' has no worksheet payload.");
            return new XlsxSheetV1(
                relationshipId,
                sheetName,
                sheetIndex,
                null,
                new Dictionary<string, XlsxCellV1>(StringComparer.OrdinalIgnoreCase));
        }

        var dimension = worksheet.SheetDimension?.Reference?.Value;
        var cells = new Dictionary<string, XlsxCellV1>(StringComparer.OrdinalIgnoreCase);

        foreach (var row in worksheet.Descendants<Row>())
        {
            var rowIndex = row.RowIndex?.Value ?? 0;
            var fallbackColumn = 1;

            foreach (var cell in row.Elements<Cell>())
            {
                var address = cell.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(address))
                {
                    if (rowIndex == 0)
                    {
                        warnings.Add($"A cell in sheet '{sheetName}' has no address and no row index; it was skipped.");
                        continue;
                    }

                    address = $"{ColumnNameFromIndex(fallbackColumn)}{rowIndex}";
                }

                var parsedCell = ParseCell(cell, address, rowIndex, fallbackColumn, sharedStrings, sheetName, warnings);
                cells[parsedCell.Address] = parsedCell;
                fallbackColumn = parsedCell.Column + 1;
            }
        }

        return new XlsxSheetV1(
            relationshipId,
            sheetName,
            sheetIndex,
            dimension,
            cells);
    }

    private static XlsxCellV1 ParseCell(
        Cell cell,
        string address,
        uint rowIndex,
        int fallbackColumn,
        SharedStringTable? sharedStrings,
        string sheetName,
        List<string> warnings)
    {
        var row = rowIndex == 0 ? RowNumberFromAddress(address) : rowIndex;
        var column = ColumnNumberFromAddress(address) ?? fallbackColumn;
        var rawValue = cell.CellValue?.Text;
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.SharedString)
        {
            var value = ReadSharedString(rawValue, sharedStrings, address, sheetName, warnings);
            return new XlsxCellV1(address, row, column, XlsxCellValueKind.String, rawValue, value, value);
        }

        if (dataType == CellValues.InlineString)
        {
            var value = cell.InlineString?.InnerText ?? string.Empty;
            return new XlsxCellV1(address, row, column, XlsxCellValueKind.String, value, value, value);
        }

        if (dataType == CellValues.String)
        {
            var value = rawValue ?? string.Empty;
            return new XlsxCellV1(address, row, column, XlsxCellValueKind.String, rawValue, value, value);
        }

        if (dataType == CellValues.Boolean)
        {
            var value = rawValue == "1" || rawValue?.Equals("true", StringComparison.OrdinalIgnoreCase) == true;
            return new XlsxCellV1(address, row, column, XlsxCellValueKind.Boolean, rawValue, value, value.ToString().ToLowerInvariant());
        }

        if (dataType == CellValues.Error)
        {
            var value = rawValue ?? string.Empty;
            return new XlsxCellV1(address, row, column, XlsxCellValueKind.Error, rawValue, value, value);
        }

        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return new XlsxCellV1(address, row, column, XlsxCellValueKind.Blank, rawValue, null, string.Empty);
        }

        if (decimal.TryParse(rawValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var number))
        {
            return new XlsxCellV1(address, row, column, XlsxCellValueKind.Number, rawValue, number, rawValue);
        }

        warnings.Add($"Cell '{sheetName}'!{address} looked numeric but could not be parsed: '{rawValue}'.");
        return new XlsxCellV1(address, row, column, XlsxCellValueKind.Number, rawValue, rawValue, rawValue);
    }

    private static string ReadSharedString(
        string? rawValue,
        SharedStringTable? sharedStrings,
        string address,
        string sheetName,
        List<string> warnings)
    {
        if (!int.TryParse(rawValue, NumberStyles.None, CultureInfo.InvariantCulture, out var sharedStringIndex))
        {
            warnings.Add($"Cell '{sheetName}'!{address} has an invalid shared string index: '{rawValue}'.");
            return rawValue ?? string.Empty;
        }

        var value = sharedStrings?.Elements<SharedStringItem>().ElementAtOrDefault(sharedStringIndex)?.InnerText;
        if (value is null)
        {
            warnings.Add($"Cell '{sheetName}'!{address} points to missing shared string index {sharedStringIndex}.");
            return rawValue ?? string.Empty;
        }

        return value;
    }

    private static uint RowNumberFromAddress(string address)
    {
        var digits = new string(address.Where(char.IsDigit).ToArray());
        return uint.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var row) ? row : 0;
    }

    private static int? ColumnNumberFromAddress(string address)
    {
        var letters = new string(address.TakeWhile(char.IsLetter).ToArray());
        if (letters.Length == 0)
        {
            return null;
        }

        var column = 0;
        foreach (var letter in letters.ToUpperInvariant())
        {
            column *= 26;
            column += letter - 'A' + 1;
        }

        return column;
    }

    private static string ColumnNameFromIndex(int columnIndex)
    {
        var columnName = new StringBuilder();
        while (columnIndex > 0)
        {
            columnIndex--;
            columnName.Insert(0, (char)('A' + columnIndex % 26));
            columnIndex /= 26;
        }

        return columnName.ToString();
    }
}
