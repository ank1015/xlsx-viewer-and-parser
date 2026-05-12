using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Parsers.Xlsx;

namespace OpenXml.Parsers.Tests;

public class XlsxWorkbookV2ParserTests
{
    [Fact]
    public void Parse_ReturnsFormulaTextAndCachedResults()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateFormulaWorkbook(path);

            var result = XlsxWorkbookV2Parser.Parse(path);

            Assert.Equal(XlsxWorkbookParseResultV2.CurrentSchemaVersion, result.SchemaVersion);
            Assert.Empty(result.Warnings);

            var sheet = Assert.Single(result.Workbook.Sheets);
            Assert.Equal("Formulas", sheet.Name);
            Assert.Equal("A1:F2", sheet.Dimension);

            var numericFormula = sheet.Cells["C1"];
            Assert.Equal(XlsxCellValueKind.Number, numericFormula.Type);
            Assert.Equal(5.5m, numericFormula.Value);
            Assert.Equal("5.5", numericFormula.RawValue);
            Assert.NotNull(numericFormula.Formula);
            Assert.Equal("A1+B1", numericFormula.Formula.Text);
            Assert.Equal("A1+B1", numericFormula.Formula.ResolvedText);
            Assert.Equal(XlsxFormulaKind.Normal, numericFormula.Formula.Kind);
            Assert.Null(numericFormula.Formula.BaseAddress);
            Assert.Null(numericFormula.Formula.BaseText);
            Assert.Equal("5.5", numericFormula.Formula.CachedRawValue);
            Assert.Equal(5.5m, numericFormula.Formula.CachedValue);
            Assert.Equal(XlsxCellValueKind.Number, numericFormula.Formula.CachedValueType);

            var stringFormula = sheet.Cells["D1"];
            Assert.Equal("Done", stringFormula.Value);
            Assert.Equal(XlsxCellValueKind.String, stringFormula.Type);
            Assert.Equal("IF(C1>5,\"Done\",\"Wait\")", stringFormula.Formula?.Text);
            Assert.Equal("Done", stringFormula.Formula?.CachedValue);

            var booleanFormula = sheet.Cells["E1"];
            Assert.Equal(true, booleanFormula.Value);
            Assert.Equal(XlsxCellValueKind.Boolean, booleanFormula.Type);
            Assert.Equal("true", booleanFormula.DisplayValue);
            Assert.Equal(XlsxCellValueKind.Boolean, booleanFormula.Formula?.CachedValueType);

            var errorFormula = sheet.Cells["F1"];
            Assert.Equal("#DIV/0!", errorFormula.Value);
            Assert.Equal(XlsxCellValueKind.Error, errorFormula.Type);
            Assert.Equal("#DIV/0!", errorFormula.Formula?.CachedValue);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void Parse_ReturnsSharedFormulaMetadata()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateSharedFormulaWorkbook(path);

            var result = XlsxWorkbookV2Parser.Parse(path);

            var sheet = Assert.Single(result.Workbook.Sheets);

            var master = sheet.Cells["C1"];
            Assert.Equal(3m, master.Value);
            Assert.NotNull(master.Formula);
            Assert.Equal("A1+$B$1", master.Formula.Text);
            Assert.Equal("A1+$B$1", master.Formula.ResolvedText);
            Assert.Equal(XlsxFormulaKind.Shared, master.Formula.Kind);
            Assert.Equal(0U, master.Formula.SharedIndex);
            Assert.Equal("C1:C2", master.Formula.Reference);
            Assert.Equal("C1", master.Formula.BaseAddress);
            Assert.Equal("A1+$B$1", master.Formula.BaseText);

            var dependent = sheet.Cells["C2"];
            Assert.Equal(7m, dependent.Value);
            Assert.NotNull(dependent.Formula);
            Assert.Null(dependent.Formula.Text);
            Assert.Equal("A2+$B$1", dependent.Formula.ResolvedText);
            Assert.Equal(XlsxFormulaKind.Shared, dependent.Formula.Kind);
            Assert.Equal(0U, dependent.Formula.SharedIndex);
            Assert.Null(dependent.Formula.Reference);
            Assert.Equal("C1", dependent.Formula.BaseAddress);
            Assert.Equal("A1+$B$1", dependent.Formula.BaseText);
            Assert.Equal(7m, dependent.Formula.CachedValue);
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

    private static void CreateFormulaWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:F2" },
            new SheetData(
                new Row(
                    NumberCell("A1", "2"),
                    NumberCell("B1", "3.5"),
                    FormulaCell("C1", "A1+B1", "5.5"),
                    FormulaCell("D1", "IF(C1>5,\"Done\",\"Wait\")", "Done", CellValues.String),
                    FormulaCell("E1", "C1>5", "1", CellValues.Boolean),
                    FormulaCell("F1", "1/0", "#DIV/0!", CellValues.Error))
                {
                    RowIndex = 1
                }));

        AddSheet(workbookPart, worksheetPart, "Formulas");
        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static void CreateSharedFormulaWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:C2" },
            new SheetData(
                new Row(
                    NumberCell("A1", "1"),
                    NumberCell("B1", "2"),
                    SharedFormulaMasterCell("C1", "A1+$B$1", "3"))
                {
                    RowIndex = 1
                },
                new Row(
                    NumberCell("A2", "3"),
                    NumberCell("B2", "4"),
                    SharedFormulaDependentCell("C2", "7"))
                {
                    RowIndex = 2
                }));

        AddSheet(workbookPart, worksheetPart, "Shared Formulas");
        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static void AddSheet(WorkbookPart workbookPart, WorksheetPart worksheetPart, string name)
    {
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook was not initialized.");
        var sheets = workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = name
        });
    }

    private static Cell NumberCell(string address, string value)
    {
        return new Cell
        {
            CellReference = address,
            CellValue = new CellValue(value)
        };
    }

    private static Cell FormulaCell(string address, string formulaText, string cachedValue, CellValues? dataType = null)
    {
        var cell = new Cell
        {
            CellReference = address,
            CellFormula = new CellFormula(formulaText),
            CellValue = new CellValue(cachedValue)
        };

        if (dataType is not null)
        {
            cell.DataType = dataType.Value;
        }

        return cell;
    }

    private static Cell SharedFormulaMasterCell(string address, string formulaText, string cachedValue)
    {
        return new Cell
        {
            CellReference = address,
            CellFormula = new CellFormula(formulaText)
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = 0U,
                Reference = "C1:C2"
            },
            CellValue = new CellValue(cachedValue)
        };
    }

    private static Cell SharedFormulaDependentCell(string address, string cachedValue)
    {
        return new Cell
        {
            CellReference = address,
            CellFormula = new CellFormula
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = 0U
            },
            CellValue = new CellValue(cachedValue)
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
