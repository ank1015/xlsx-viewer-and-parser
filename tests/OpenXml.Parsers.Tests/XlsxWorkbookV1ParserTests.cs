using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Parsers.Xlsx;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace OpenXml.Parsers.Tests;

public class XlsxWorkbookV1ParserTests
{
    [Fact]
    public void Parse_ReturnsCellsV1WorkbookShape()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateCellsWorkbook(path);

            var result = XlsxWorkbookV1Parser.Parse(path);

            Assert.Equal(XlsxWorkbookParseResultV1.CurrentSchemaVersion, result.SchemaVersion);
            Assert.Empty(result.Warnings);
            Assert.Equal(Path.GetFileName(path), result.Source.FileName);

            var sheet = Assert.Single(result.Workbook.Sheets);
            Assert.Equal("Cells", sheet.Name);
            Assert.Equal(0, sheet.Index);
            Assert.Equal("A1:G1", sheet.Dimension);

            Assert.Equal("Shared Name", sheet.Cells["A1"].Value);
            Assert.Equal(XlsxCellValueKind.String, sheet.Cells["A1"].Type);
            Assert.Equal("0", sheet.Cells["A1"].RawValue);

            Assert.Equal("Inline Label", sheet.Cells["B1"].Value);
            Assert.Equal(XlsxCellValueKind.String, sheet.Cells["B1"].Type);

            Assert.Equal("Plain Text", sheet.Cells["C1"].Value);
            Assert.Equal(XlsxCellValueKind.String, sheet.Cells["C1"].Type);

            Assert.Equal(42.5m, sheet.Cells["D1"].Value);
            Assert.Equal(XlsxCellValueKind.Number, sheet.Cells["D1"].Type);
            Assert.Equal("42.5", sheet.Cells["D1"].DisplayValue);

            Assert.Equal(true, sheet.Cells["E1"].Value);
            Assert.Equal(XlsxCellValueKind.Boolean, sheet.Cells["E1"].Type);
            Assert.Equal("true", sheet.Cells["E1"].DisplayValue);

            Assert.Null(sheet.Cells["F1"].Value);
            Assert.Equal(XlsxCellValueKind.Blank, sheet.Cells["F1"].Type);
            Assert.Equal(string.Empty, sheet.Cells["F1"].DisplayValue);

            Assert.Equal("#DIV/0!", sheet.Cells["G1"].Value);
            Assert.Equal(XlsxCellValueKind.Error, sheet.Cells["G1"].Type);
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

    private static void CreateCellsWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new SharedStringTable(
            new SharedStringItem(new SpreadsheetText("Shared Name")));
        sharedStringPart.SharedStringTable.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:G1" },
            new SheetData(
                new Row(
                    SharedStringCell("A1", 0),
                    InlineStringCell("B1", "Inline Label"),
                    PlainStringCell("C1", "Plain Text"),
                    NumberCell("D1", "42.5"),
                    BooleanCell("E1", true),
                    BlankCell("F1"),
                    ErrorCell("G1", "#DIV/0!"))
                {
                    RowIndex = 1
                }));

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Cells"
        });

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static Cell SharedStringCell(string address, int sharedStringIndex)
    {
        return new Cell
        {
            CellReference = address,
            DataType = CellValues.SharedString,
            CellValue = new CellValue(sharedStringIndex.ToString())
        };
    }

    private static Cell InlineStringCell(string address, string value)
    {
        return new Cell
        {
            CellReference = address,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new SpreadsheetText(value))
        };
    }

    private static Cell PlainStringCell(string address, string value)
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

    private static Cell BooleanCell(string address, bool value)
    {
        return new Cell
        {
            CellReference = address,
            DataType = CellValues.Boolean,
            CellValue = new CellValue(value ? "1" : "0")
        };
    }

    private static Cell BlankCell(string address)
    {
        return new Cell
        {
            CellReference = address
        };
    }

    private static Cell ErrorCell(string address, string value)
    {
        return new Cell
        {
            CellReference = address,
            DataType = CellValues.Error,
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
