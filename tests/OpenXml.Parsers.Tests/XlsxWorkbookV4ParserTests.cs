using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Parsers.Xlsx;

namespace OpenXml.Parsers.Tests;

public class XlsxWorkbookV4ParserTests
{
    [Fact]
    public void Parse_ReturnsMergedCellsRowsAndColumns()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateLayoutWorkbook(path);

            var result = XlsxWorkbookV4Parser.Parse(path);

            Assert.Equal(XlsxWorkbookParseResultV4.CurrentSchemaVersion, result.SchemaVersion);
            Assert.Empty(result.Warnings);

            var sheet = Assert.Single(result.Workbook.Sheets);
            Assert.Equal("Layout", sheet.Name);
            Assert.Equal("A1:D4", sheet.Dimension);

            var merge = Assert.Single(sheet.MergedCells);
            Assert.Equal("A1:C2", merge.Reference);
            Assert.Equal("A1", merge.StartAddress);
            Assert.Equal("C2", merge.EndAddress);
            Assert.Equal(1U, merge.StartRow);
            Assert.Equal(1, merge.StartColumn);
            Assert.Equal(2U, merge.EndRow);
            Assert.Equal(3, merge.EndColumn);

            Assert.Equal(2, sheet.Columns.Count);
            var wideColumns = sheet.Columns[0];
            Assert.Equal(1U, wideColumns.Min);
            Assert.Equal(3U, wideColumns.Max);
            Assert.Equal(18.5, wideColumns.Width);
            Assert.True(wideColumns.CustomWidth);
            Assert.False(wideColumns.Hidden);

            var hiddenColumn = sheet.Columns[1];
            Assert.Equal(4U, hiddenColumn.Min);
            Assert.Equal(4U, hiddenColumn.Max);
            Assert.True(hiddenColumn.Hidden);
            Assert.Equal(1U, hiddenColumn.StyleIndex);

            Assert.Equal(3, sheet.Rows.Count);
            var tallRow = sheet.Rows.Single(row => row.Index == 1);
            Assert.Equal(24.5, tallRow.Height);
            Assert.True(tallRow.CustomHeight);
            Assert.Equal(1U, tallRow.StyleIndex);

            var hiddenRow = sheet.Rows.Single(row => row.Index == 4);
            Assert.True(hiddenRow.Hidden);
            Assert.Equal((byte)2, hiddenRow.OutlineLevel);

            Assert.Equal("Merged title", sheet.Cells["A1"].Value);
            Assert.Equal(42m, sheet.Cells["D4"].Value);
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

    private static void CreateLayoutWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(new CellFormat(), new CellFormat { ApplyFont = true }) { Count = 2 });
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:D4" },
            new Columns(
                new Column
                {
                    Min = 1,
                    Max = 3,
                    Width = 18.5,
                    CustomWidth = true
                },
                new Column
                {
                    Min = 4,
                    Max = 4,
                    Width = 0,
                    Hidden = true,
                    Style = 1
                }),
            new SheetData(
                new Row(
                    TextCell("A1", "Merged title"))
                {
                    RowIndex = 1,
                    Height = 24.5,
                    CustomHeight = true,
                    StyleIndex = 1
                },
                new Row
                {
                    RowIndex = 2
                },
                new Row(
                    NumberCell("D4", "42"))
                {
                    RowIndex = 4,
                    Hidden = true,
                    OutlineLevel = 2
                }),
            new MergeCells(
                new MergeCell { Reference = "A1:C2" })
            {
                Count = 1
            });

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Layout"
        });

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static Cell TextCell(string address, string value)
    {
        return new Cell
        {
            CellReference = address,
            DataType = CellValues.String,
            CellValue = new CellValue(value)
        };
    }

    private static Cell NumberCell(string address, string value)
    {
        return new Cell
        {
            CellReference = address,
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
