using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Parsers.Xlsx;

namespace OpenXml.Parsers.Tests;

public class XlsxWorkbookV3ParserTests
{
    [Fact]
    public void Parse_ReturnsStylesAndFormattedCellValues()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateStyledWorkbook(path);

            var result = XlsxWorkbookV3Parser.Parse(path);

            Assert.Equal(XlsxWorkbookParseResultV3.CurrentSchemaVersion, result.SchemaVersion);
            Assert.Empty(result.Warnings);

            var styles = result.Workbook.Styles;
            Assert.Equal(4, styles.CellFormats.Count);
            Assert.Contains(styles.NumberFormats, format => format.Id == 14 && format.IsBuiltIn && format.IsDateFormat);
            Assert.Contains(styles.NumberFormats, format => format.Id == 165 && !format.IsBuiltIn && format.FormatCode == "$#,##0.00");

            var customFont = styles.Fonts[1];
            Assert.Equal("Aptos", customFont.Name);
            Assert.Equal(14, customFont.Size);
            Assert.True(customFont.Bold);
            Assert.True(customFont.Italic);
            Assert.Equal("FFFF0000", customFont.Color?.Rgb);
            Assert.Equal("FF0000", customFont.Color?.ResolvedRgb);

            var fill = styles.Fills[2];
            Assert.Equal("solid", fill.PatternType);
            Assert.Equal("FFFFFF00", fill.ForegroundColor?.Rgb);
            Assert.Equal("FFFF00", fill.ForegroundColor?.ResolvedRgb);

            var border = styles.Borders[1];
            Assert.Equal("thin", border.Left.Style);
            Assert.Equal("FF0000FF", border.Left.Color?.Rgb);
            Assert.Equal("0000FF", border.Left.Color?.ResolvedRgb);

            var sheet = Assert.Single(result.Workbook.Sheets);

            var dateCell = sheet.Cells["A1"];
            Assert.Equal(XlsxCellValueKind.Date, dateCell.Type);
            Assert.Equal("1/15/24", dateCell.DisplayValue);
            Assert.Equal(1U, dateCell.StyleIndex);
            Assert.Equal(14U, dateCell.Style?.NumberFormatId);
            Assert.True(dateCell.Style?.IsDateFormat);
            var dateValue = Assert.IsType<DateTime>(dateCell.Value);
            Assert.Equal(new DateTime(2024, 1, 15), dateValue);

            var currencyCell = sheet.Cells["B1"];
            Assert.Equal(XlsxCellValueKind.Number, currencyCell.Type);
            Assert.Equal(1234.5m, currencyCell.Value);
            Assert.Equal("$1,234.50", currencyCell.DisplayValue);
            Assert.Equal(2U, currencyCell.StyleIndex);
            Assert.Equal(165U, currencyCell.Style?.NumberFormatId);
            Assert.Equal(1U, currencyCell.Style?.FontId);
            Assert.Equal(2U, currencyCell.Style?.FillId);
            Assert.Equal(1U, currencyCell.Style?.BorderId);
            Assert.True(currencyCell.Style?.ApplyAlignment);
            Assert.True(currencyCell.Style?.WrapText);

            var decimalCell = sheet.Cells["C1"];
            Assert.Equal(XlsxCellValueKind.Number, decimalCell.Type);
            Assert.Equal("9.8", decimalCell.DisplayValue);
            Assert.Equal(3U, decimalCell.StyleIndex);
            Assert.Equal("0.0", decimalCell.Style?.NumberFormatCode);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void Parse_ReturnsThemeColorsAndResolvesThemeTintColors()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateThemeStyledWorkbook(path);

            var result = XlsxWorkbookV3Parser.Parse(path);

            Assert.Empty(result.Warnings);
            var styles = result.Workbook.Styles;
            var accent5 = Assert.Single(styles.ThemeColors, color => color.Index == 8);
            Assert.Equal("accent5", accent5.Name);
            Assert.Equal("A02B93", accent5.Rgb);

            var themeFontColor = styles.Fonts[1].Color;
            Assert.Equal(8U, themeFontColor?.Theme);
            Assert.Equal(0.5d, themeFontColor?.Tint);
            Assert.Null(themeFontColor?.Rgb);
            Assert.Equal("D095C9", themeFontColor?.ResolvedRgb);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    private static string CreateTempPath(string extension)
    {
        return Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}{extension}");
    }

    private static void CreateStyledWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStylesheet();
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:C1" },
            new SheetData(
                new Row(
                    NumberCell("A1", new DateTime(2024, 1, 15).ToOADate().ToString(CultureInfo.InvariantCulture), 1),
                    NumberCell("B1", "1234.5", 2),
                    NumberCell("C1", "9.8000000000000007", 3))
                {
                    RowIndex = 1
                }));

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Styles"
        });

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static void CreateThemeStyledWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        AddThemePart(workbookPart);

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateThemeStylesheet();
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:A1" },
            new SheetData(
                new Row(NumberCell("A1", "1", 1))
                {
                    RowIndex = 1
                }));

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Theme"
        });

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static void AddThemePart(WorkbookPart workbookPart)
    {
        var themePart = workbookPart.AddNewPart<ThemePart>();
        var themeXml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
              <a:themeElements>
                <a:clrScheme name="Office">
                  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
                  <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
                  <a:dk2><a:srgbClr val="0E2841"/></a:dk2>
                  <a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>
                  <a:accent1><a:srgbClr val="156082"/></a:accent1>
                  <a:accent2><a:srgbClr val="E97132"/></a:accent2>
                  <a:accent3><a:srgbClr val="196B24"/></a:accent3>
                  <a:accent4><a:srgbClr val="0F9ED5"/></a:accent4>
                  <a:accent5><a:srgbClr val="A02B93"/></a:accent5>
                  <a:accent6><a:srgbClr val="4EA72E"/></a:accent6>
                  <a:hlink><a:srgbClr val="467886"/></a:hlink>
                  <a:folHlink><a:srgbClr val="96607D"/></a:folHlink>
                </a:clrScheme>
              </a:themeElements>
            </a:theme>
            """;

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(themeXml));
        themePart.FeedData(stream);
    }

    private static Stylesheet CreateStylesheet()
    {
        return new Stylesheet(
            new NumberingFormats(
                new NumberingFormat
                {
                    NumberFormatId = 165,
                    FormatCode = "$#,##0.00"
                },
                new NumberingFormat
                {
                    NumberFormatId = 166,
                    FormatCode = "0.0"
                })
            {
                Count = 2
            },
            new Fonts(
                new Font(
                    new FontSize { Val = 11 },
                    new FontName { Val = "Calibri" }),
                new Font(
                    new Bold(),
                    new Italic(),
                    new FontSize { Val = 14 },
                    new Color { Rgb = "FFFF0000" },
                    new FontName { Val = "Aptos" }))
            {
                Count = 2
            },
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                new Fill(
                    new PatternFill(
                        new ForegroundColor { Rgb = "FFFFFF00" },
                        new BackgroundColor { Indexed = 64 })
                    {
                        PatternType = PatternValues.Solid
                    }))
            {
                Count = 3
            },
            new Borders(
                new Border(),
                new Border(
                    new LeftBorder(new Color { Rgb = "FF0000FF" }) { Style = BorderStyleValues.Thin },
                    new RightBorder { Style = BorderStyleValues.Thin },
                    new TopBorder { Style = BorderStyleValues.Thin },
                    new BottomBorder { Style = BorderStyleValues.Thin },
                    new DiagonalBorder()))
            {
                Count = 2
            },
            new CellStyleFormats(new CellFormat())
            {
                Count = 1
            },
            new CellFormats(
                new CellFormat(),
                new CellFormat
                {
                    NumberFormatId = 14,
                    ApplyNumberFormat = true
                },
                new CellFormat
                {
                    NumberFormatId = 165,
                    FontId = 1,
                    FillId = 2,
                    BorderId = 1,
                    ApplyNumberFormat = true,
                    ApplyFont = true,
                    ApplyFill = true,
                    ApplyBorder = true,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                new CellFormat
                {
                    NumberFormatId = 166,
                    ApplyNumberFormat = true
                })
            {
                Count = 4
            });
    }

    private static Stylesheet CreateThemeStylesheet()
    {
        return new Stylesheet(
            new Fonts(
                new Font(
                    new FontSize { Val = 11 },
                    new FontName { Val = "Calibri" }),
                new Font(
                    new Color { Theme = 8U, Tint = 0.5d },
                    new FontName { Val = "Aptos" }))
            {
                Count = 2
            },
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
            {
                Count = 2
            },
            new Borders(new Border())
            {
                Count = 1
            },
            new CellStyleFormats(new CellFormat())
            {
                Count = 1
            },
            new CellFormats(
                new CellFormat(),
                new CellFormat
                {
                    FontId = 1,
                    ApplyFont = true
                })
            {
                Count = 2
            });
    }

    private static Cell NumberCell(string address, string value, uint styleIndex)
    {
        return new Cell
        {
            CellReference = address,
            StyleIndex = styleIndex,
            CellValue = new CellValue(value)
        };
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
