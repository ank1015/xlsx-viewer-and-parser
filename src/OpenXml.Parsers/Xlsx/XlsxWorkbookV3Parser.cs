using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXml.Parsers.Xlsx;

public static partial class XlsxWorkbookV3Parser
{
    public static XlsxWorkbookParseResultV3 Parse(string filePath)
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
        var styles = ParseStyles(workbookPart);
        var date1904 = workbookPart.Workbook?.WorkbookProperties?.Date1904?.Value == true;
        var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
        var parsedSheets = new List<XlsxSheetV3>();

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
                styles,
                date1904,
                warnings));
        }

        return new XlsxWorkbookParseResultV3(
            XlsxWorkbookParseResultV3.CurrentSchemaVersion,
            new XlsxSourceInfoV1(fileInfo.Name, fileInfo.FullName, fileInfo.Length),
            new XlsxWorkbookV3(parsedSheets, styles),
            warnings);
    }

    private static XlsxSheetV3 ParseSheet(
        WorksheetPart worksheetPart,
        string relationshipId,
        string sheetName,
        int sheetIndex,
        SharedStringTable? sharedStrings,
        XlsxStylesV3 styles,
        bool date1904,
        List<string> warnings)
    {
        var worksheet = worksheetPart.Worksheet;
        if (worksheet is null)
        {
            warnings.Add($"Sheet '{sheetName}' has no worksheet payload.");
            return new XlsxSheetV3(
                relationshipId,
                sheetName,
                sheetIndex,
                null,
                new Dictionary<string, XlsxCellV3>(StringComparer.OrdinalIgnoreCase));
        }

        var dimension = worksheet.SheetDimension?.Reference?.Value;
        var cells = new Dictionary<string, XlsxCellV3>(StringComparer.OrdinalIgnoreCase);
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

                var parsedCell = ParseCell(
                    cell,
                    address,
                    rowIndex,
                    fallbackColumn,
                    sharedStrings,
                    styles,
                    date1904,
                    sharedFormulaMasters,
                    sheetName,
                    warnings);

                cells[parsedCell.Address] = parsedCell;
                fallbackColumn = parsedCell.Column + 1;
            }
        }

        return new XlsxSheetV3(
            relationshipId,
            sheetName,
            sheetIndex,
            dimension,
            cells);
    }

    private static XlsxCellV3 ParseCell(
        Cell cell,
        string address,
        uint rowIndex,
        int fallbackColumn,
        SharedStringTable? sharedStrings,
        XlsxStylesV3 styles,
        bool date1904,
        IReadOnlyDictionary<uint, SharedFormulaMaster> sharedFormulaMasters,
        string sheetName,
        List<string> warnings)
    {
        var row = rowIndex == 0 ? RowNumberFromAddress(address) : rowIndex;
        var column = ColumnNumberFromAddress(address) ?? fallbackColumn;
        var styleIndex = cell.StyleIndex?.Value;
        var style = styleIndex is null || styleIndex.Value >= styles.CellFormats.Count
            ? null
            : styles.CellFormats[(int)styleIndex.Value];
        var parsedValue = ParseCellValue(cell, address, sharedStrings, style, date1904, sheetName, warnings);
        var formula = ParseFormula(cell, address, parsedValue, sharedFormulaMasters, sheetName, warnings);

        return new XlsxCellV3(
            address,
            row,
            column,
            parsedValue.Type,
            parsedValue.RawValue,
            parsedValue.Value,
            parsedValue.DisplayValue,
            styleIndex,
            style,
            formula);
    }

    private static XlsxStylesV3 ParseStyles(WorkbookPart workbookPart)
    {
        var themeColors = ParseThemeColors(workbookPart);
        var themeColorsByIndex = themeColors.ToDictionary(color => color.Index, color => color.Rgb);
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        if (stylesheet is null)
        {
            return new XlsxStylesV3([], themeColors, [], [], [], []);
        }

        var numberFormatsById = BuiltInNumberFormats()
            .ToDictionary(
                item => item.Key,
                item => new XlsxNumberFormatV3(item.Key, item.Value, true, IsDateFormatCode(item.Value)));

        foreach (var numberFormat in stylesheet.NumberingFormats?.Elements<NumberingFormat>() ?? [])
        {
            if (numberFormat.NumberFormatId?.Value is not { } id || numberFormat.FormatCode?.Value is not { } code)
            {
                continue;
            }

            numberFormatsById[id] = new XlsxNumberFormatV3(id, code, false, IsDateFormatCode(code));
        }

        var fonts = stylesheet.Fonts?.Elements<Font>()
            .Select((font, index) => new XlsxFontV3(
                (uint)index,
                font.FontName?.Val?.Value,
                font.FontSize?.Val?.Value,
                font.Bold is not null,
                font.Italic is not null,
                font.Underline is not null,
                font.Strike is not null,
                ParseColor(font.Color, themeColorsByIndex)))
            .ToList() ?? [];

        var fills = stylesheet.Fills?.Elements<Fill>()
            .Select((fill, index) =>
            {
                var patternFill = fill.PatternFill;
                return new XlsxFillV3(
                    (uint)index,
                    patternFill?.PatternType?.InnerText,
                    ParseColor(patternFill?.ForegroundColor, themeColorsByIndex),
                    ParseColor(patternFill?.BackgroundColor, themeColorsByIndex));
            })
            .ToList() ?? [];

        var borders = stylesheet.Borders?.Elements<Border>()
            .Select((border, index) => new XlsxBorderV3(
                (uint)index,
                ParseBorderSide(border.LeftBorder, themeColorsByIndex),
                ParseBorderSide(border.RightBorder, themeColorsByIndex),
                ParseBorderSide(border.TopBorder, themeColorsByIndex),
                ParseBorderSide(border.BottomBorder, themeColorsByIndex),
                ParseBorderSide(border.DiagonalBorder, themeColorsByIndex)))
            .ToList() ?? [];

        var cellFormats = stylesheet.CellFormats?.Elements<CellFormat>()
            .Select((cellFormat, index) =>
            {
                var numberFormatId = cellFormat.NumberFormatId?.Value;
                numberFormatsById.TryGetValue(numberFormatId ?? 0, out var numberFormat);
                return new XlsxCellFormatV3(
                    (uint)index,
                    numberFormatId,
                    numberFormat?.FormatCode,
                    numberFormat?.IsDateFormat == true,
                    cellFormat.FontId?.Value,
                    cellFormat.FillId?.Value,
                    cellFormat.BorderId?.Value,
                    cellFormat.ApplyNumberFormat?.Value == true,
                    cellFormat.ApplyFont?.Value == true,
                    cellFormat.ApplyFill?.Value == true,
                    cellFormat.ApplyBorder?.Value == true,
                    cellFormat.ApplyAlignment?.Value == true,
                    cellFormat.Alignment?.WrapText?.Value == true);
            })
            .ToList() ?? [];

        return new XlsxStylesV3(
            numberFormatsById.Values.OrderBy(format => format.Id).ToList(),
            themeColors,
            fonts,
            fills,
            borders,
            cellFormats);
    }

    private static IReadOnlyList<XlsxThemeColorV3> ParseThemeColors(WorkbookPart workbookPart)
    {
        var themePart = workbookPart.GetPartsOfType<ThemePart>().FirstOrDefault();
        if (themePart is null)
        {
            return [];
        }

        using var stream = themePart.GetStream();
        var document = XDocument.Load(stream);
        var colorScheme = document.Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "clrScheme");
        if (colorScheme is null)
        {
            return [];
        }

        var colorsByName = colorScheme.Elements()
            .Select(element => new
            {
                Name = element.Name.LocalName,
                Rgb = ReadThemeColorRgb(element)
            })
            .Where(color => color.Rgb is not null)
            .ToDictionary(
                color => color.Name,
                color => color.Rgb!,
                StringComparer.OrdinalIgnoreCase);

        var themeOrder = new (uint Index, string Name)[]
        {
            (0, "lt1"),
            (1, "dk1"),
            (2, "lt2"),
            (3, "dk2"),
            (4, "accent1"),
            (5, "accent2"),
            (6, "accent3"),
            (7, "accent4"),
            (8, "accent5"),
            (9, "accent6"),
            (10, "hlink"),
            (11, "folHlink")
        };

        return themeOrder
            .Where(item => colorsByName.ContainsKey(item.Name))
            .Select(item => new XlsxThemeColorV3(item.Index, item.Name, colorsByName[item.Name]))
            .ToList();
    }

    private static string? ReadThemeColorRgb(XElement themeColor)
    {
        var srgb = themeColor.Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "srgbClr")
            ?.Attribute("val")
            ?.Value;
        if (!string.IsNullOrWhiteSpace(srgb))
        {
            return NormalizeRgb(srgb);
        }

        var systemColor = themeColor.Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "sysClr");
        return NormalizeRgb(systemColor?.Attribute("lastClr")?.Value);
    }

    private static XlsxColorV3? ParseColor(
        ColorType? color,
        IReadOnlyDictionary<uint, string> themeColorsByIndex)
    {
        if (color is null)
        {
            return null;
        }

        var rgb = color.Rgb?.Value;
        var resolvedRgb = ResolveColorRgb(rgb, color.Indexed?.Value, color.Theme?.Value, color.Tint?.Value, themeColorsByIndex);

        return new XlsxColorV3(
            rgb,
            color.Indexed?.Value,
            color.Theme?.Value,
            color.Tint?.Value,
            color.Auto?.Value == true,
            resolvedRgb);
    }

    private static string? ResolveColorRgb(
        string? rgb,
        uint? indexed,
        uint? theme,
        double? tint,
        IReadOnlyDictionary<uint, string> themeColorsByIndex)
    {
        var baseRgb = default(string);

        if (theme is not null && themeColorsByIndex.TryGetValue(theme.Value, out var themeRgb))
        {
            baseRgb = themeRgb;
        }
        else if (!string.IsNullOrWhiteSpace(rgb))
        {
            baseRgb = NormalizeRgb(rgb);
        }
        else if (indexed is not null)
        {
            baseRgb = ResolveIndexedColor(indexed.Value);
        }

        if (baseRgb is null)
        {
            return null;
        }

        return ApplyTint(baseRgb, tint);
    }

    private static string? ResolveIndexedColor(uint indexed)
    {
        return indexed switch
        {
            0 => "000000",
            1 => "FFFFFF",
            2 => "FF0000",
            3 => "00FF00",
            4 => "0000FF",
            5 => "FFFF00",
            6 => "FF00FF",
            7 => "00FFFF",
            8 => "000000",
            9 => "FFFFFF",
            10 => "FF0000",
            11 => "00FF00",
            12 => "0000FF",
            13 => "FFFF00",
            14 => "FF00FF",
            15 => "00FFFF",
            16 => "800000",
            17 => "008000",
            18 => "000080",
            19 => "808000",
            20 => "800080",
            21 => "008080",
            22 => "C0C0C0",
            23 => "808080",
            24 => "9999FF",
            25 => "993366",
            26 => "FFFFCC",
            27 => "CCFFFF",
            28 => "660066",
            29 => "FF8080",
            30 => "0066CC",
            31 => "CCCCFF",
            32 => "000080",
            33 => "FF00FF",
            34 => "FFFF00",
            35 => "00FFFF",
            36 => "800080",
            37 => "800000",
            38 => "008080",
            39 => "0000FF",
            40 => "00CCFF",
            41 => "CCFFFF",
            42 => "CCFFCC",
            43 => "FFFF99",
            44 => "99CCFF",
            45 => "FF99CC",
            46 => "CC99FF",
            47 => "FFCC99",
            48 => "3366FF",
            49 => "33CCCC",
            50 => "99CC00",
            51 => "FFCC00",
            52 => "FF9900",
            53 => "FF6600",
            54 => "666699",
            55 => "969696",
            56 => "003366",
            57 => "339966",
            58 => "003300",
            59 => "333300",
            60 => "993300",
            61 => "993366",
            62 => "333399",
            63 => "333333",
            _ => null
        };
    }

    private static string? NormalizeRgb(string? rgb)
    {
        if (string.IsNullOrWhiteSpace(rgb))
        {
            return null;
        }

        var normalized = rgb.Trim().TrimStart('#');
        if (normalized.Length == 8)
        {
            normalized = normalized[2..];
        }

        return normalized.Length == 6 ? normalized.ToUpperInvariant() : null;
    }

    private static string ApplyTint(string rgb, double? tint)
    {
        if (tint is null || Math.Abs(tint.Value) < double.Epsilon)
        {
            return rgb;
        }

        var normalizedTint = Math.Clamp(tint.Value, -1d, 1d);
        var red = ApplyTintToChannel(Convert.ToInt32(rgb[0..2], 16), normalizedTint);
        var green = ApplyTintToChannel(Convert.ToInt32(rgb[2..4], 16), normalizedTint);
        var blue = ApplyTintToChannel(Convert.ToInt32(rgb[4..6], 16), normalizedTint);

        return $"{red:X2}{green:X2}{blue:X2}";
    }

    private static int ApplyTintToChannel(int channel, double tint)
    {
        var tinted = tint < 0
            ? channel * (1d + tint)
            : channel * (1d - tint) + 255d * tint;

        return Math.Clamp((int)Math.Round(tinted, MidpointRounding.AwayFromZero), 0, 255);
    }

    private static XlsxBorderSideV3 ParseBorderSide(
        BorderPropertiesType? side,
        IReadOnlyDictionary<uint, string> themeColorsByIndex)
    {
        return new XlsxBorderSideV3(
            side?.Style?.InnerText,
            ParseColor(side?.Color, themeColorsByIndex));
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
        XlsxCellFormatV3? style,
        bool date1904,
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

        if (!decimal.TryParse(rawValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var number))
        {
            warnings.Add($"Cell '{sheetName}'!{address} looked numeric but could not be parsed: '{rawValue}'.");
            return new ParsedCellValue(XlsxCellValueKind.Number, rawValue, rawValue, rawValue);
        }

        if (style?.IsDateFormat == true &&
            double.TryParse(rawValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var serialDate))
        {
            var date = ReadExcelDate(serialDate, date1904);
            return new ParsedCellValue(
                XlsxCellValueKind.Date,
                rawValue,
                date,
                FormatDate(date, style.NumberFormatCode));
        }

        return new ParsedCellValue(
            XlsxCellValueKind.Number,
            rawValue,
            number,
            FormatNumber(number, style?.NumberFormatCode, rawValue));
    }

    private static DateTime ReadExcelDate(double serialDate, bool date1904)
    {
        return date1904
            ? new DateTime(1904, 1, 1).AddDays(serialDate)
            : DateTime.FromOADate(serialDate);
    }

    private static string FormatDate(DateTime date, string? formatCode)
    {
        var normalized = NormalizeFormatCode(formatCode);
        if (normalized.Contains("yyyy-mm-dd", StringComparison.OrdinalIgnoreCase))
        {
            return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        if (normalized.Contains("mm-dd-yy", StringComparison.OrdinalIgnoreCase))
        {
            return date.ToString("MM-dd-yy", CultureInfo.InvariantCulture);
        }

        if (normalized.Contains("m/d/yy", StringComparison.OrdinalIgnoreCase))
        {
            return date.ToString("M/d/yy", CultureInfo.InvariantCulture);
        }

        if (ContainsTimeToken(normalized) && ContainsDateToken(normalized))
        {
            return date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        if (ContainsTimeToken(normalized))
        {
            return date.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
        }

        return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    private static string FormatNumber(decimal number, string? formatCode, string rawValue)
    {
        var normalized = NormalizeFormatCode(formatCode);
        if (string.IsNullOrWhiteSpace(normalized) || normalized.Equals("general", StringComparison.OrdinalIgnoreCase))
        {
            return rawValue;
        }

        if (normalized.Contains('%'))
        {
            return (number * 100m).ToString("0.##", CultureInfo.InvariantCulture) + "%";
        }

        if (normalized.Contains('$'))
        {
            return "$" + number.ToString("#,##0.00", CultureInfo.InvariantCulture);
        }

        if (normalized.Contains("#,##0.00", StringComparison.OrdinalIgnoreCase))
        {
            return number.ToString("#,##0.00", CultureInfo.InvariantCulture);
        }

        if (DecimalPlacesFromFormat(normalized) is { } decimalPlaces)
        {
            return number.ToString($"0.{new string('0', decimalPlaces)}", CultureInfo.InvariantCulture);
        }

        if (normalized.Contains("#,##0", StringComparison.OrdinalIgnoreCase))
        {
            return number.ToString("#,##0", CultureInfo.InvariantCulture);
        }

        return rawValue;
    }

    private static int? DecimalPlacesFromFormat(string normalizedFormatCode)
    {
        var match = DecimalFormatRegex().Match(normalizedFormatCode);

        return match.Success ? match.Groups["places"].Value.Length : null;
    }

    private static string NormalizeFormatCode(string? formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode))
        {
            return string.Empty;
        }

        var withoutSections = formatCode.Split(';')[0];
        return FormatLiteralRegex().Replace(FormatConditionRegex().Replace(withoutSections, string.Empty), string.Empty);
    }

    private static bool IsDateFormatCode(string formatCode)
    {
        var normalized = NormalizeFormatCode(formatCode);
        return ContainsDateToken(normalized) || ContainsTimeToken(normalized);
    }

    private static bool ContainsDateToken(string formatCode)
    {
        return formatCode.Contains('y', StringComparison.OrdinalIgnoreCase) ||
               formatCode.Contains('d', StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsTimeToken(string formatCode)
    {
        return formatCode.Contains('h', StringComparison.OrdinalIgnoreCase) ||
               formatCode.Contains('s', StringComparison.OrdinalIgnoreCase);
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

    private static IReadOnlyDictionary<uint, string> BuiltInNumberFormats()
    {
        return new Dictionary<uint, string>
        {
            [0] = "General",
            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [4] = "#,##0.00",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00",
            [12] = "# ?/?",
            [13] = "# ??/??",
            [14] = "m/d/yy",
            [15] = "d-mmm-yy",
            [16] = "d-mmm",
            [17] = "mmm-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "m/d/yy h:mm",
            [37] = "#,##0 ;(#,##0)",
            [38] = "#,##0 ;[Red](#,##0)",
            [39] = "#,##0.00;(#,##0.00)",
            [40] = "#,##0.00;[Red](#,##0.00)",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [48] = "##0.0E+0",
            [49] = "@"
        };
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

    [GeneratedRegex(@"\[[^\]]+\]")]
    private static partial Regex FormatConditionRegex();

    [GeneratedRegex("\"[^\"]*\"|\\\\.")]
    private static partial Regex FormatLiteralRegex();

    [GeneratedRegex(@"(?:^|[^A-Za-z])0\.(?<places>[0#]+)")]
    private static partial Regex DecimalFormatRegex();

    private sealed record ParsedCellValue(
        XlsxCellValueKind Type,
        string? RawValue,
        object? Value,
        string DisplayValue);

    private sealed record SharedFormulaMaster(string Address, string Text);
}
