using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXml.Parsers.Xlsx;

public static partial class XlsxWorkbookV2Parser
{
    public static XlsxWorkbookParseResultV2 Parse(string filePath)
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
        var parsedSheets = new List<XlsxSheetV2>();

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

        return new XlsxWorkbookParseResultV2(
            XlsxWorkbookParseResultV2.CurrentSchemaVersion,
            new XlsxSourceInfoV1(fileInfo.Name, fileInfo.FullName, fileInfo.Length),
            new XlsxWorkbookV2(parsedSheets),
            warnings);
    }

    private static XlsxSheetV2 ParseSheet(
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
            return new XlsxSheetV2(
                relationshipId,
                sheetName,
                sheetIndex,
                null,
                new Dictionary<string, XlsxCellV2>(StringComparer.OrdinalIgnoreCase));
        }

        var dimension = worksheet.SheetDimension?.Reference?.Value;
        var cells = new Dictionary<string, XlsxCellV2>(StringComparer.OrdinalIgnoreCase);
        var sharedFormulaMasters = CollectSharedFormulaMasters(worksheet);

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

                var parsedCell = ParseCell(cell, address, rowIndex, fallbackColumn, sharedStrings, sharedFormulaMasters, sheetName, warnings);
                cells[parsedCell.Address] = parsedCell;
                fallbackColumn = parsedCell.Column + 1;
            }
        }

        return new XlsxSheetV2(
            relationshipId,
            sheetName,
            sheetIndex,
            dimension,
            cells);
    }

    private static XlsxCellV2 ParseCell(
        Cell cell,
        string address,
        uint rowIndex,
        int fallbackColumn,
        SharedStringTable? sharedStrings,
        IReadOnlyDictionary<uint, SharedFormulaMaster> sharedFormulaMasters,
        string sheetName,
        List<string> warnings)
    {
        var row = rowIndex == 0 ? RowNumberFromAddress(address) : rowIndex;
        var column = ColumnNumberFromAddress(address) ?? fallbackColumn;
        var parsedValue = ParseCellValue(cell, address, sharedStrings, sheetName, warnings);
        var formula = ParseFormula(cell, address, parsedValue, sharedFormulaMasters, sheetName, warnings);

        return new XlsxCellV2(
            address,
            row,
            column,
            parsedValue.Type,
            parsedValue.RawValue,
            parsedValue.Value,
            parsedValue.DisplayValue,
            formula);
    }

    private static XlsxFormulaV2? ParseFormula(
        Cell cell,
        string address,
        ParsedCellValue parsedValue,
        IReadOnlyDictionary<uint, SharedFormulaMaster> sharedFormulaMasters,
        string sheetName,
        List<string> warnings)
    {
        var formula = cell.CellFormula;
        if (formula is null)
        {
            return null;
        }

        var text = string.IsNullOrWhiteSpace(formula.Text) ? null : formula.Text;
        var kind = ReadFormulaKind(formula);
        var sharedIndex = formula.SharedIndex?.Value;
        var baseAddress = default(string);
        var baseText = default(string);
        var resolvedText = text;

        if (kind == XlsxFormulaKind.Shared)
        {
            if (text is not null)
            {
                baseAddress = address;
                baseText = text;
            }
            else if (sharedIndex is not null && sharedFormulaMasters.TryGetValue(sharedIndex.Value, out var master))
            {
                baseAddress = master.Address;
                baseText = master.Text;
                resolvedText = TranslateSharedFormula(master.Text, master.Address, address);
            }
            else
            {
                warnings.Add($"Cell '{sheetName}'!{address} is a shared formula but its master formula was not found.");
            }
        }

        return new XlsxFormulaV2(
            text,
            resolvedText,
            kind,
            formula.SharedIndex?.Value,
            formula.Reference?.Value,
            baseAddress,
            baseText,
            parsedValue.RawValue,
            parsedValue.Value,
            parsedValue.DisplayValue,
            parsedValue.Type);
    }

    private static Dictionary<uint, SharedFormulaMaster> CollectSharedFormulaMasters(Worksheet worksheet)
    {
        var masters = new Dictionary<uint, SharedFormulaMaster>();

        foreach (var cell in worksheet.Descendants<Cell>())
        {
            var formula = cell.CellFormula;
            if (formula?.FormulaType?.Value != CellFormulaValues.Shared ||
                formula.SharedIndex?.Value is not { } sharedIndex ||
                string.IsNullOrWhiteSpace(formula.Text) ||
                string.IsNullOrWhiteSpace(cell.CellReference?.Value))
            {
                continue;
            }

            masters[sharedIndex] = new SharedFormulaMaster(cell.CellReference.Value!, formula.Text);
        }

        return masters;
    }

    private static XlsxFormulaKind ReadFormulaKind(CellFormula formula)
    {
        var formulaType = formula.FormulaType?.Value;
        if (formulaType == CellFormulaValues.Shared)
        {
            return XlsxFormulaKind.Shared;
        }

        if (formulaType == CellFormulaValues.Array)
        {
            return XlsxFormulaKind.Array;
        }

        if (formulaType == CellFormulaValues.DataTable)
        {
            return XlsxFormulaKind.DataTable;
        }

        return XlsxFormulaKind.Normal;
    }

    private static ParsedCellValue ParseCellValue(
        Cell cell,
        string address,
        SharedStringTable? sharedStrings,
        string sheetName,
        List<string> warnings)
    {
        var rawValue = cell.CellValue?.Text;
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.SharedString)
        {
            var value = ReadSharedString(rawValue, sharedStrings, address, sheetName, warnings);
            return new ParsedCellValue(XlsxCellValueKind.String, rawValue, value, value);
        }

        if (dataType == CellValues.InlineString)
        {
            var value = cell.InlineString?.InnerText ?? string.Empty;
            return new ParsedCellValue(XlsxCellValueKind.String, value, value, value);
        }

        if (dataType == CellValues.String)
        {
            var value = rawValue ?? string.Empty;
            return new ParsedCellValue(XlsxCellValueKind.String, rawValue, value, value);
        }

        if (dataType == CellValues.Boolean)
        {
            var value = rawValue == "1" || rawValue?.Equals("true", StringComparison.OrdinalIgnoreCase) == true;
            return new ParsedCellValue(XlsxCellValueKind.Boolean, rawValue, value, value.ToString().ToLowerInvariant());
        }

        if (dataType == CellValues.Error)
        {
            var value = rawValue ?? string.Empty;
            return new ParsedCellValue(XlsxCellValueKind.Error, rawValue, value, value);
        }

        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return new ParsedCellValue(XlsxCellValueKind.Blank, rawValue, null, string.Empty);
        }

        if (decimal.TryParse(rawValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var number))
        {
            return new ParsedCellValue(XlsxCellValueKind.Number, rawValue, number, rawValue);
        }

        warnings.Add($"Cell '{sheetName}'!{address} looked numeric but could not be parsed: '{rawValue}'.");
        return new ParsedCellValue(XlsxCellValueKind.Number, rawValue, rawValue, rawValue);
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

    private static string TranslateSharedFormula(string formulaText, string baseAddress, string targetAddress)
    {
        var baseRow = RowNumberFromAddress(baseAddress);
        var targetRow = RowNumberFromAddress(targetAddress);
        var baseColumn = ColumnNumberFromAddress(baseAddress);
        var targetColumn = ColumnNumberFromAddress(targetAddress);

        if (baseRow == 0 || targetRow == 0 || baseColumn is null || targetColumn is null)
        {
            return formulaText;
        }

        var rowOffset = (int)targetRow - (int)baseRow;
        var columnOffset = targetColumn.Value - baseColumn.Value;
        return TranslateFormulaReferences(formulaText, rowOffset, columnOffset);
    }

    private static string TranslateFormulaReferences(string formulaText, int rowOffset, int columnOffset)
    {
        var builder = new StringBuilder(formulaText.Length);
        var segment = new StringBuilder();
        var insideString = false;

        for (var i = 0; i < formulaText.Length; i++)
        {
            var character = formulaText[i];
            if (character == '"')
            {
                if (insideString && i + 1 < formulaText.Length && formulaText[i + 1] == '"')
                {
                    builder.Append(character);
                    builder.Append(formulaText[++i]);
                    continue;
                }

                builder.Append(character);
                if (insideString)
                {
                    insideString = false;
                    continue;
                }

                builder.Length--;
                builder.Append(TranslateFormulaSegment(segment.ToString(), rowOffset, columnOffset));
                segment.Clear();
                builder.Append(character);
                insideString = true;
                continue;
            }

            if (insideString)
            {
                builder.Append(character);
            }
            else
            {
                segment.Append(character);
            }
        }

        if (segment.Length > 0)
        {
            builder.Append(TranslateFormulaSegment(segment.ToString(), rowOffset, columnOffset));
        }

        return builder.ToString();
    }

    private static string TranslateFormulaSegment(string formulaSegment, int rowOffset, int columnOffset)
    {
        return CellReferenceRegex().Replace(formulaSegment, match =>
        {
            var columnAbsolute = match.Groups["columnAbsolute"].Value == "$";
            var rowAbsolute = match.Groups["rowAbsolute"].Value == "$";
            var columnText = match.Groups["column"].Value;
            var rowText = match.Groups["row"].Value;

            var column = ColumnNumberFromAddress(columnText);
            if (column is null || !int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out var row))
            {
                return match.Value;
            }

            var translatedColumn = columnAbsolute ? column.Value : column.Value + columnOffset;
            var translatedRow = rowAbsolute ? row : row + rowOffset;
            if (translatedColumn < 1 || translatedRow < 1)
            {
                return match.Value;
            }

            return string.Concat(
                match.Groups["sheet"].Value,
                columnAbsolute ? "$" : string.Empty,
                ColumnNameFromIndex(translatedColumn),
                rowAbsolute ? "$" : string.Empty,
                translatedRow.ToString(CultureInfo.InvariantCulture));
        });
    }

    [GeneratedRegex(@"(?<![A-Za-z0-9_])(?<sheet>(?:'[^']+'|[A-Za-z_][A-Za-z0-9_\.]*)!)?(?<columnAbsolute>\$?)(?<column>[A-Za-z]{1,3})(?<rowAbsolute>\$?)(?<row>\d{1,7})(?![A-Za-z0-9_])")]
    private static partial Regex CellReferenceRegex();

    private sealed record ParsedCellValue(
        XlsxCellValueKind Type,
        string? RawValue,
        object? Value,
        string DisplayValue);

    private sealed record SharedFormulaMaster(string Address, string Text);
}
