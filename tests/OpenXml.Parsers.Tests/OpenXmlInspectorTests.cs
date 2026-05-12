using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Parsers;
using WordBody = DocumentFormat.OpenXml.Wordprocessing.Body;
using WordDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using WordParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WordRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace OpenXml.Parsers.Tests;

public class OpenXmlInspectorTests
{
    [Fact]
    public void Inspect_ReadsDocxParagraphPreview()
    {
        var path = CreateTempPath(".docx");
        try
        {
            CreateDocx(path, "Hello OpenXML");

            var result = OpenXmlInspector.Inspect(path);

            Assert.Equal(OpenXmlFileKind.WordDocument, result.Kind);
            Assert.Equal("1", result.Summary["paragraphs"]);
            Assert.Contains("Hello OpenXML", result.PreviewRows);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void Inspect_ReadsXlsxSheetNamesAndPreviewRows()
    {
        var path = CreateTempPath(".xlsx");
        try
        {
            CreateXlsx(path);

            var result = OpenXmlInspector.Inspect(path);

            Assert.Equal(OpenXmlFileKind.ExcelWorkbook, result.Kind);
            Assert.Equal("1", result.Summary["sheets"]);
            Assert.Equal("Items", result.Summary["sheetNames"]);
            Assert.Contains("Name | Count", result.PreviewRows);
            Assert.Contains("Parser | 3", result.PreviewRows);
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

    private static void CreateDocx(string path, string paragraphText)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new WordDocument(
            new WordBody(
                new WordParagraph(
                    new WordRun(
                        new WordText(paragraphText)))));

        mainPart.Document.Save();
    }

    private static void CreateXlsx(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetData(
                new Row(
                    TextCell("Name"),
                    TextCell("Count")),
                new Row(
                    TextCell("Parser"),
                    NumberCell("3"))));

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Items"
        });

        workbookPart.Workbook.Save();
    }

    private static Cell TextCell(string value)
    {
        return new Cell
        {
            DataType = CellValues.String,
            CellValue = new CellValue(value)
        };
    }

    private static Cell NumberCell(string value)
    {
        return new Cell
        {
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
