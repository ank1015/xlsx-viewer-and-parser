using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using OpenXml.Parsers.Xlsx;

namespace OpenXml.Parsers.Tests;

public class XlsxWorkbookV6ParserTests
{
    [Fact]
    public void Parse_ReturnsDrawingChartAndImageMetadata()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateDrawingWorkbook(path);

            var result = XlsxWorkbookV6Parser.Parse(path);

            Assert.Equal(XlsxWorkbookParseResultV6.CurrentSchemaVersion, result.SchemaVersion);
            Assert.Empty(result.Warnings);

            var sheet = Assert.Single(result.Workbook.Sheets);
            var drawing = Assert.Single(sheet.Drawings);
            Assert.Equal("rIdDrawing1", drawing.RelationshipId);

            var chart = Assert.Single(sheet.Charts);
            Assert.Equal("chart-1", chart.Id);
            Assert.Equal("rIdChart1", chart.RelationshipId);
            Assert.Equal("bar", chart.Type);
            Assert.Equal("col", chart.BarDirection);
            Assert.Equal("clustered", chart.BarGrouping);
            Assert.Equal("Sales Chart", chart.Title);
            Assert.Null(chart.AssetPath);
            Assert.Equal("twoCell", chart.Anchor?.Kind);
            Assert.Equal(1, chart.Anchor?.From?.Column);
            Assert.Equal(1U, chart.Anchor?.From?.Row);
            Assert.Equal(5, chart.Anchor?.To?.Column);
            Assert.Equal(10U, chart.Anchor?.To?.Row);

            var series = Assert.Single(chart.Series);
            Assert.Equal("Revenue", series.Name);
            Assert.Equal("Sheet1!$A$2:$A$3", series.CategoriesRange);
            Assert.Equal("Sheet1!$B$2:$B$3", series.ValuesRange);
            Assert.Equal(new[] { "Jan", "Feb" }, series.CachedCategories);
            Assert.Equal(new[] { 10d, 20d }, series.CachedValues);

            var image = Assert.Single(sheet.Images);
            Assert.Equal("image-1", image.Id);
            Assert.Equal("rIdImage1", image.RelationshipId);
            Assert.Equal("Logo", image.Name);
            Assert.Equal("Tiny logo", image.Description);
            Assert.Equal("image/png", image.ContentType);
            Assert.Equal("image-1.png", image.FileName);
            Assert.Null(image.AssetPath);
            Assert.Equal("oneCell", image.Anchor?.Kind);
            Assert.Equal(6, image.Anchor?.From?.Column);
            Assert.Equal(1U, image.Anchor?.From?.Row);
            Assert.Equal(914400L, image.Anchor?.CxEmu);
            Assert.Equal(914400L, image.Anchor?.CyEmu);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void Parse_ExtractsChartXmlAndImagesWhenAssetDirectoryIsProvided()
    {
        var path = CreateTempPath(".xlsx");
        var outputDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        try
        {
            CreateDrawingWorkbook(path);

            var result = XlsxWorkbookV6Parser.Parse(path, outputDirectory);

            var sheet = Assert.Single(result.Workbook.Sheets);
            var chart = Assert.Single(sheet.Charts);
            var image = Assert.Single(sheet.Images);

            Assert.Equal("assets/charts/chart-1.xml", chart.AssetPath);
            Assert.Equal("assets/images/image-1.png", image.AssetPath);
            var chartAssetPath = Assert.IsType<string>(chart.AssetPath);
            var imageAssetPath = Assert.IsType<string>(image.AssetPath);
            Assert.True(File.Exists(Path.Combine(outputDirectory, chartAssetPath)));
            Assert.True(File.Exists(Path.Combine(outputDirectory, imageAssetPath)));
            Assert.Contains("barChart", File.ReadAllText(Path.Combine(outputDirectory, chartAssetPath)));
            Assert.NotEqual(0, new FileInfo(Path.Combine(outputDirectory, imageAssetPath)).Length);
        }
        finally
        {
            DeleteIfExists(path);
            if (Directory.Exists(outputDirectory))
            {
                Directory.Delete(outputDirectory, true);
            }
        }
    }

    private static string CreateTempPath(string extension)
    {
        return Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}{extension}");
    }

    private static void CreateDrawingWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:B3" },
            new SheetData(
                new Row(
                    TextCell("A1", "Month"),
                    TextCell("B1", "Revenue"))
                {
                    RowIndex = 1
                },
                new Row(
                    TextCell("A2", "Jan"),
                    NumberCell("B2", "10"))
                {
                    RowIndex = 2
                },
                new Row(
                    TextCell("A3", "Feb"),
                    NumberCell("B3", "20"))
                {
                    RowIndex = 3
                }));

        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>("rIdDrawing1");
        var imagePart = drawingsPart.AddImagePart(ImagePartType.Png, "rIdImage1");
        using (var stream = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=")))
        {
            imagePart.FeedData(stream);
        }

        var chartPart = drawingsPart.AddNewPart<ChartPart>("rIdChart1");
        chartPart.ChartSpace = CreateChartSpace();
        chartPart.ChartSpace.Save();

        drawingsPart.WorksheetDrawing = CreateWorksheetDrawing();
        drawingsPart.WorksheetDrawing.Save();

        worksheetPart.Worksheet.Append(new Drawing { Id = "rIdDrawing1" });

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static Xdr.WorksheetDrawing CreateWorksheetDrawing()
    {
        return new Xdr.WorksheetDrawing(
            new Xdr.TwoCellAnchor(
                FromMarker(0, 0),
                ToMarker(4, 9),
                new Xdr.GraphicFrame(
                    new Xdr.NonVisualGraphicFrameProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Sales Chart" },
                        new Xdr.NonVisualGraphicFrameDrawingProperties()),
                    new Xdr.Transform(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = 0, Cy = 0 }),
                    new A.Graphic(
                        new A.GraphicData(
                            new C.ChartReference { Id = "rIdChart1" })
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                        })),
                new Xdr.ClientData()),
            new Xdr.OneCellAnchor(
                FromMarker(5, 0),
                new Xdr.Extent { Cx = 914400L, Cy = 914400L },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties
                        {
                            Id = 3U,
                            Name = "Logo",
                            Description = "Tiny logo"
                        },
                        new Xdr.NonVisualPictureDrawingProperties()),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = "rIdImage1" },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = 914400L, Cy = 914400L }),
                        new A.PresetGeometry(new A.AdjustValueList())
                        {
                            Preset = A.ShapeTypeValues.Rectangle
                        })),
                new Xdr.ClientData()));
    }

    private static Xdr.FromMarker FromMarker(int column, int row)
    {
        return new Xdr.FromMarker(
            new Xdr.ColumnId(column.ToString()),
            new Xdr.ColumnOffset("0"),
            new Xdr.RowId(row.ToString()),
            new Xdr.RowOffset("0"));
    }

    private static Xdr.ToMarker ToMarker(int column, int row)
    {
        return new Xdr.ToMarker(
            new Xdr.ColumnId(column.ToString()),
            new Xdr.ColumnOffset("0"),
            new Xdr.RowId(row.ToString()),
            new Xdr.RowOffset("0"));
    }

    private static C.ChartSpace CreateChartSpace()
    {
        return new C.ChartSpace(
            new C.Chart(
                new C.Title(
                    new C.ChartText(
                        new C.RichText(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(
                                new A.Run(
                                    new A.Text("Sales Chart")))))),
                new C.PlotArea(
                    new C.Layout(),
                    new C.BarChart(
                        new C.BarDirection { Val = C.BarDirectionValues.Column },
                        new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                        new C.BarChartSeries(
                            new C.Index { Val = 0U },
                            new C.Order { Val = 0U },
                            new C.SeriesText(
                                new C.StringReference(
                                    new C.Formula("Sheet1!$B$1"),
                                    new C.StringCache(
                                        new C.PointCount { Val = 1U },
                                        new C.StringPoint(
                                            new C.NumericValue("Revenue"))
                                        {
                                            Index = 0U
                                        }))),
                            new C.CategoryAxisData(
                                new C.StringReference(
                                    new C.Formula("Sheet1!$A$2:$A$3"),
                                    new C.StringCache(
                                        new C.PointCount { Val = 2U },
                                        new C.StringPoint(new C.NumericValue("Jan")) { Index = 0U },
                                        new C.StringPoint(new C.NumericValue("Feb")) { Index = 1U }))),
                            new C.Values(
                                new C.NumberReference(
                                    new C.Formula("Sheet1!$B$2:$B$3"),
                                    new C.NumberingCache(
                                        new C.FormatCode("General"),
                                        new C.PointCount { Val = 2U },
                                        new C.NumericPoint(new C.NumericValue("10")) { Index = 0U },
                                        new C.NumericPoint(new C.NumericValue("20")) { Index = 1U }))))))));
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
