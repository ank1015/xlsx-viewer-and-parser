using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Parsers.Xlsx;

namespace OpenXml.Parsers.Tests;

public class XlsxWorkbookV5ParserTests
{
    [Fact]
    public void Parse_ReturnsTableDefinitions()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateTableWorkbook(path);

            var result = XlsxWorkbookV5Parser.Parse(path);

            Assert.Equal(XlsxWorkbookParseResultV5.CurrentSchemaVersion, result.SchemaVersion);
            Assert.Empty(result.Warnings);

            var sheet = Assert.Single(result.Workbook.Sheets);
            Assert.Equal("Sales", sheet.Name);

            var table = Assert.Single(sheet.Tables);
            Assert.Equal("rIdTable1", table.RelationshipId);
            Assert.Equal(1U, table.Id);
            Assert.Equal("SalesTable", table.Name);
            Assert.Equal("SalesTable", table.DisplayName);
            Assert.Equal("A1:C4", table.Reference);
            Assert.Equal("A1", table.StartAddress);
            Assert.Equal("C4", table.EndAddress);
            Assert.Equal(1U, table.StartRow);
            Assert.Equal(1, table.StartColumn);
            Assert.Equal(4U, table.EndRow);
            Assert.Equal(3, table.EndColumn);
            Assert.Equal(1U, table.HeaderRowCount);
            Assert.Equal(1U, table.TotalsRowCount);
            Assert.True(table.Published);

            Assert.NotNull(table.AutoFilter);
            Assert.Equal("A1:C3", table.AutoFilter.Reference);
            var filterColumn = Assert.Single(table.AutoFilter.FilterColumns);
            Assert.Equal(1U, filterColumn.ColumnId);
            Assert.True(filterColumn.ShowButton);
            Assert.Contains("filters", filterColumn.FilterTypes);

            Assert.Equal(3, table.Columns.Count);
            Assert.Equal(new[] { "Product", "Quantity", "Total" }, table.Columns.Select(column => column.Name));
            Assert.Equal("Total", table.Columns[0].TotalsRowLabel);
            Assert.Equal("sum", table.Columns[1].TotalsRowFunction);
            Assert.Equal("SUM([Quantity])", table.Columns[1].TotalsRowFormula);
            Assert.Equal("[Quantity]*10", table.Columns[2].CalculatedColumnFormula);
            Assert.Equal("sum", table.Columns[2].TotalsRowFunction);

            Assert.NotNull(table.Style);
            Assert.Equal("TableStyleMedium2", table.Style.Name);
            Assert.True(table.Style.ShowRowStripes);
            Assert.False(table.Style.ShowColumnStripes);

            Assert.Equal("Widget", sheet.Cells["A2"].Value);
            Assert.Equal(4m, sheet.Cells["B4"].Value);
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

    private static void CreateTableWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:C4" },
            new SheetData(
                new Row(
                    TextCell("A1", "Product"),
                    TextCell("B1", "Quantity"),
                    TextCell("C1", "Total"))
                {
                    RowIndex = 1
                },
                new Row(
                    TextCell("A2", "Widget"),
                    NumberCell("B2", "2"),
                    NumberCell("C2", "20"))
                {
                    RowIndex = 2
                },
                new Row(
                    TextCell("A3", "Gadget"),
                    NumberCell("B3", "2"),
                    NumberCell("C3", "20"))
                {
                    RowIndex = 3
                },
                new Row(
                    TextCell("A4", "Total"),
                    NumberCell("B4", "4"),
                    NumberCell("C4", "40"))
                {
                    RowIndex = 4
                }));

        var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>("rIdTable1");
        tablePart.Table = CreateTable();
        tablePart.Table.Save();

        worksheetPart.Worksheet.Append(new TableParts(
            new TablePart { Id = "rIdTable1" })
        {
            Count = 1
        });

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sales"
        });

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static Table CreateTable()
    {
        return new Table(
            new AutoFilter(
                new FilterColumn(
                    new Filters(
                        new Filter { Val = "2" }))
                {
                    ColumnId = 1,
                    ShowButton = true
                })
            {
                Reference = "A1:C3"
            },
            new TableColumns(
                new TableColumn
                {
                    Id = 1,
                    Name = "Product",
                    TotalsRowLabel = "Total"
                },
                new TableColumn(
                    new TotalsRowFormula("SUM([Quantity])"))
                {
                    Id = 2,
                    Name = "Quantity",
                    TotalsRowFunction = TotalsRowFunctionValues.Sum
                },
                new TableColumn(
                    new CalculatedColumnFormula("[Quantity]*10"))
                {
                    Id = 3,
                    Name = "Total",
                    TotalsRowFunction = TotalsRowFunctionValues.Sum
                })
            {
                Count = 3
            },
            new TableStyleInfo
            {
                Name = "TableStyleMedium2",
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            })
        {
            Id = 1,
            Name = "SalesTable",
            DisplayName = "SalesTable",
            Reference = "A1:C4",
            HeaderRowCount = 1,
            TotalsRowCount = 1,
            Published = true
        };
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
