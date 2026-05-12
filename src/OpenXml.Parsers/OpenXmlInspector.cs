using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WordParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace OpenXml.Parsers;

public static class OpenXmlInspector
{
    public static OpenXmlInspectionResult Inspect(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("A file path is required.", nameof(filePath));
        }

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("OpenXML file was not found.", filePath);
        }

        return Path.GetExtension(filePath).ToLowerInvariant() switch
        {
            ".docx" => InspectWordDocument(filePath),
            ".xlsx" => InspectWorkbook(filePath),
            ".pptx" => InspectPresentation(filePath),
            _ => new OpenXmlInspectionResult(
                filePath,
                OpenXmlFileKind.Unknown,
                new Dictionary<string, string>
                {
                    ["extension"] = Path.GetExtension(filePath),
                    ["status"] = "Unsupported OpenXML extension."
                },
                [])
        };
    }

    private static OpenXmlInspectionResult InspectWordDocument(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var body = document.MainDocumentPart?.Document?.Body;
        var paragraphs = body?.Descendants<WordParagraph>().ToList() ?? [];
        var preview = paragraphs
            .Select(p => string.Join("", p.Descendants<WordText>().Select(t => t.Text)).Trim())
            .Where(text => text.Length > 0)
            .Take(10)
            .ToArray();

        return new OpenXmlInspectionResult(
            filePath,
            OpenXmlFileKind.WordDocument,
            new Dictionary<string, string>
            {
                ["paragraphs"] = paragraphs.Count.ToString(),
                ["previewItems"] = preview.Length.ToString()
            },
            preview);
    }

    private static OpenXmlInspectionResult InspectWorkbook(string filePath)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart;
        var sheets = workbookPart?.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
        var preview = new List<string>();

        var firstSheet = sheets.FirstOrDefault();
        if (workbookPart is not null && firstSheet?.Id?.Value is { } firstSheetId)
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheetId);
            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

            foreach (var row in worksheetPart.Worksheet?.Descendants<Row>().Take(10) ?? [])
            {
                var cells = row
                    .Elements<Cell>()
                    .Take(8)
                    .Select(cell => ReadCellText(cell, sharedStrings))
                    .ToArray();

                if (cells.Any(value => !string.IsNullOrWhiteSpace(value)))
                {
                    preview.Add(string.Join(" | ", cells));
                }
            }
        }

        return new OpenXmlInspectionResult(
            filePath,
            OpenXmlFileKind.ExcelWorkbook,
            new Dictionary<string, string>
            {
                ["sheets"] = sheets.Count.ToString(),
                ["sheetNames"] = string.Join(", ", sheets.Select(sheet => sheet.Name?.Value).Where(name => !string.IsNullOrWhiteSpace(name))),
                ["previewItems"] = preview.Count.ToString()
            },
            preview);
    }

    private static OpenXmlInspectionResult InspectPresentation(string filePath)
    {
        using var document = PresentationDocument.Open(filePath, false);
        var slideCount = document.PresentationPart?.SlideParts.Count() ?? 0;

        return new OpenXmlInspectionResult(
            filePath,
            OpenXmlFileKind.PowerPointPresentation,
            new Dictionary<string, string>
            {
                ["slides"] = slideCount.ToString(),
                ["status"] = "Presentation package opened successfully."
            },
            []);
    }

    private static string ReadCellText(Cell cell, SharedStringTable? sharedStrings)
    {
        var rawValue = cell.CellValue?.Text ?? string.Empty;
        if (cell.DataType?.Value == CellValues.SharedString && int.TryParse(rawValue, out var sharedStringIndex))
        {
            return sharedStrings?.ElementAtOrDefault(sharedStringIndex)?.InnerText ?? rawValue;
        }

        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? string.Empty;
        }

        return rawValue;
    }
}
