using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OpenXml.Parsers.Xlsx;

public static class XlsxWorkbookV6Parser
{
    public static XlsxWorkbookParseResultV6 Parse(string filePath, string? assetOutputDirectory = null)
    {
        var v5 = XlsxWorkbookV5Parser.Parse(filePath);
        var warnings = v5.Warnings.ToList();

        if (!string.IsNullOrWhiteSpace(assetOutputDirectory))
        {
            Directory.CreateDirectory(Path.Combine(assetOutputDirectory, "assets", "images"));
            Directory.CreateDirectory(Path.Combine(assetOutputDirectory, "assets", "charts"));
        }

        var drawingContext = new DrawingParseContext(
            assetOutputDirectory,
            v5.Workbook.Styles.ThemeColors.ToDictionary(
                color => color.Name,
                color => color.Rgb,
                StringComparer.OrdinalIgnoreCase));
        var sheetDrawings = ParseSheetDrawings(filePath, drawingContext, warnings);

        var sheets = v5.Workbook.Sheets
            .Select(sheet =>
            {
                sheetDrawings.TryGetValue(sheet.Id, out var drawings);
                drawings ??= [];

                return new XlsxSheetV6(
                    sheet.Id,
                    sheet.Name,
                    sheet.Index,
                    sheet.Dimension,
                    sheet.Cells,
                    sheet.MergedCells,
                    sheet.Rows,
                    sheet.Columns,
                    sheet.Tables,
                    drawings,
                    drawings.SelectMany(drawing => drawing.Charts).ToList(),
                    drawings.SelectMany(drawing => drawing.Images).ToList());
            })
            .ToList();

        return new XlsxWorkbookParseResultV6(
            XlsxWorkbookParseResultV6.CurrentSchemaVersion,
            v5.Source,
            new XlsxWorkbookV6(sheets, v5.Workbook.Styles),
            warnings);
    }

    private static Dictionary<string, IReadOnlyList<XlsxDrawingV6>> ParseSheetDrawings(
        string filePath,
        DrawingParseContext context,
        List<string> warnings)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("Workbook part is missing.");

        var drawingsBySheetId = new Dictionary<string, IReadOnlyList<XlsxDrawingV6>>(StringComparer.OrdinalIgnoreCase);
        var sheets = workbookPart.Workbook?.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToList() ?? [];

        foreach (var sheet in sheets)
        {
            var relationshipId = sheet.Id?.Value;
            var sheetName = sheet.Name?.Value ?? relationshipId ?? "Unknown";
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            if (workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var drawings = new List<XlsxDrawingV6>();
            var drawingPart = worksheetPart.DrawingsPart;
            if (drawingPart is null)
            {
                drawingsBySheetId[relationshipId] = drawings;
                continue;
            }

            var drawingRelationshipId = worksheetPart.GetIdOfPart(drawingPart);
            var parsed = ParseDrawingPart(drawingPart, drawingRelationshipId, sheetName, context, warnings);
            drawings.Add(parsed);
            drawingsBySheetId[relationshipId] = drawings;
        }

        return drawingsBySheetId;
    }

    private static XlsxDrawingV6 ParseDrawingPart(
        DrawingsPart drawingPart,
        string drawingRelationshipId,
        string sheetName,
        DrawingParseContext context,
        List<string> warnings)
    {
        var charts = new List<XlsxChartV6>();
        var images = new List<XlsxImageV6>();
        var worksheetDrawing = drawingPart.WorksheetDrawing;

        if (worksheetDrawing is null)
        {
            warnings.Add($"Sheet '{sheetName}' has drawing relationship '{drawingRelationshipId}' with no drawing payload.");
            return new XlsxDrawingV6(drawingRelationshipId, charts, images);
        }

        foreach (var anchor in worksheetDrawing.ChildElements)
        {
            var parsedAnchor = ParseAnchor(anchor);

            foreach (var graphicFrame in anchor.Descendants<Xdr.GraphicFrame>())
            {
                var chartReference = graphicFrame.Descendants<C.ChartReference>().FirstOrDefault();
                var chartRelationshipId = chartReference?.Id?.Value;
                if (string.IsNullOrWhiteSpace(chartRelationshipId))
                {
                    continue;
                }

                if (drawingPart.GetPartById(chartRelationshipId) is not ChartPart chartPart)
                {
                    warnings.Add($"Sheet '{sheetName}' drawing '{drawingRelationshipId}' references chart relationship '{chartRelationshipId}' that was not found.");
                    continue;
                }

                charts.Add(ParseChart(chartPart, chartRelationshipId, parsedAnchor, context));
            }

            foreach (var picture in anchor.Descendants<Xdr.Picture>())
            {
                var blip = picture.BlipFill?.Blip;
                var imageRelationshipId = blip?.Embed?.Value ?? blip?.Link?.Value;
                if (string.IsNullOrWhiteSpace(imageRelationshipId))
                {
                    continue;
                }

                if (drawingPart.GetPartById(imageRelationshipId) is not ImagePart imagePart)
                {
                    warnings.Add($"Sheet '{sheetName}' drawing '{drawingRelationshipId}' references image relationship '{imageRelationshipId}' that was not found.");
                    continue;
                }

                images.Add(ParseImage(imagePart, imageRelationshipId, picture, parsedAnchor, context));
            }
        }

        return new XlsxDrawingV6(drawingRelationshipId, charts, images);
    }

    private static XlsxChartV6 ParseChart(
        ChartPart chartPart,
        string relationshipId,
        XlsxDrawingAnchorV6? anchor,
        DrawingParseContext context)
    {
        var id = $"chart-{++context.ChartCount}";
        var assetPath = default(string);
        var type = ReadChartType(chartPart.ChartSpace);
        var barDirection = ReadBarDirection(chartPart.ChartSpace);
        var barGrouping = ReadBarGrouping(chartPart.ChartSpace);
        var title = ReadChartTitle(chartPart.ChartSpace);
        var series = ReadChartSeries(chartPart.ChartSpace, context);

        if (context.AssetOutputDirectory is not null)
        {
            assetPath = $"assets/charts/{id}.xml";
            var fullPath = Path.Combine(context.AssetOutputDirectory, assetPath);
            File.WriteAllText(fullPath, chartPart.ChartSpace?.OuterXml ?? string.Empty);
        }

        return new XlsxChartV6(
            id,
            relationshipId,
            type,
            barDirection,
            barGrouping,
            title,
            anchor,
            series,
            assetPath);
    }

    private static XlsxImageV6 ParseImage(
        ImagePart imagePart,
        string relationshipId,
        Xdr.Picture picture,
        XlsxDrawingAnchorV6? anchor,
        DrawingParseContext context)
    {
        var id = $"image-{++context.ImageCount}";
        var extension = ExtensionFromContentType(imagePart.ContentType);
        var fileName = $"{id}{extension}";
        var assetPath = default(string);

        if (context.AssetOutputDirectory is not null)
        {
            assetPath = $"assets/images/{fileName}";
            var fullPath = Path.Combine(context.AssetOutputDirectory, assetPath);
            using var input = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using var output = File.Create(fullPath);
            input.CopyTo(output);
        }

        var nonVisual = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
        return new XlsxImageV6(
            id,
            relationshipId,
            nonVisual?.Name?.Value,
            nonVisual?.Description?.Value,
            imagePart.ContentType,
            fileName,
            assetPath,
            anchor);
    }

    private static XlsxDrawingAnchorV6? ParseAnchor(OpenXmlElement anchor)
    {
        if (anchor is Xdr.TwoCellAnchor twoCellAnchor)
        {
            return new XlsxDrawingAnchorV6(
                "twoCell",
                ParseMarker(twoCellAnchor.FromMarker),
                ParseMarker(twoCellAnchor.ToMarker),
                null,
                null,
                null,
                null);
        }

        if (anchor is Xdr.OneCellAnchor oneCellAnchor)
        {
            return new XlsxDrawingAnchorV6(
                "oneCell",
                ParseMarker(oneCellAnchor.FromMarker),
                null,
                null,
                null,
                oneCellAnchor.Extent?.Cx?.Value,
                oneCellAnchor.Extent?.Cy?.Value);
        }

        if (anchor is Xdr.AbsoluteAnchor absoluteAnchor)
        {
            return new XlsxDrawingAnchorV6(
                "absolute",
                null,
                null,
                absoluteAnchor.Position?.X?.Value,
                absoluteAnchor.Position?.Y?.Value,
                absoluteAnchor.Extent?.Cx?.Value,
                absoluteAnchor.Extent?.Cy?.Value);
        }

        return null;
    }

    private static XlsxDrawingAnchorMarkerV6? ParseMarker(Xdr.MarkerType? marker)
    {
        if (marker is null)
        {
            return null;
        }

        return new XlsxDrawingAnchorMarkerV6(
            ParseInt(marker.ColumnId?.Text) + 1,
            ParseLong(marker.ColumnOffset?.Text),
            (uint)Math.Max(0, ParseInt(marker.RowId?.Text) + 1),
            ParseLong(marker.RowOffset?.Text));
    }

    private static string? ReadChartType(C.ChartSpace? chartSpace)
    {
        var plotArea = chartSpace?.GetFirstChild<C.Chart>()?.PlotArea;
        if (plotArea is null)
        {
            return null;
        }

        if (plotArea.Elements<C.BarChart>().Any()) return "bar";
        if (plotArea.Elements<C.LineChart>().Any()) return "line";
        if (plotArea.Elements<C.PieChart>().Any()) return "pie";
        if (plotArea.Elements<C.ScatterChart>().Any()) return "scatter";
        if (plotArea.Elements<C.AreaChart>().Any()) return "area";
        if (plotArea.Elements<C.DoughnutChart>().Any()) return "doughnut";
        if (plotArea.Elements<C.BubbleChart>().Any()) return "bubble";
        if (plotArea.Elements<C.RadarChart>().Any()) return "radar";
        if (plotArea.Elements<C.StockChart>().Any()) return "stock";
        if (plotArea.Elements<C.SurfaceChart>().Any()) return "surface";

        return plotArea.ChildElements.FirstOrDefault(element => element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase))?.LocalName;
    }

    private static string? ReadBarDirection(C.ChartSpace? chartSpace)
    {
        var direction = chartSpace
            ?.GetFirstChild<C.Chart>()
            ?.PlotArea
            ?.Elements<C.BarChart>()
            .FirstOrDefault()
            ?.BarDirection
            ?.Val
            ?.Value;

        if (direction == C.BarDirectionValues.Bar)
        {
            return "bar";
        }

        if (direction == C.BarDirectionValues.Column)
        {
            return "col";
        }

        return null;
    }

    private static string? ReadBarGrouping(C.ChartSpace? chartSpace)
    {
        var grouping = chartSpace
            ?.GetFirstChild<C.Chart>()
            ?.PlotArea
            ?.Elements<C.BarChart>()
            .FirstOrDefault()
            ?.GetFirstChild<C.BarGrouping>()
            ?.Val
            ?.Value;

        if (grouping == C.BarGroupingValues.PercentStacked)
        {
            return "percentStacked";
        }

        if (grouping == C.BarGroupingValues.Stacked)
        {
            return "stacked";
        }

        if (grouping == C.BarGroupingValues.Clustered)
        {
            return "clustered";
        }

        if (grouping == C.BarGroupingValues.Standard)
        {
            return "standard";
        }

        return null;
    }

    private static string? ReadChartTitle(C.ChartSpace? chartSpace)
    {
        var title = chartSpace?.GetFirstChild<C.Chart>()?.Title;
        if (title is null)
        {
            return null;
        }

        var text = string.Concat(title.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(item => item.Text));
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private static IReadOnlyList<XlsxChartSeriesV6> ReadChartSeries(C.ChartSpace? chartSpace, DrawingParseContext context)
    {
        var series = chartSpace?.GetFirstChild<C.Chart>()?.PlotArea?.Descendants()
            .Where(element => element.LocalName.Equals("ser", StringComparison.OrdinalIgnoreCase))
            .ToList() ?? [];

        return series
            .Select(item => new XlsxChartSeriesV6(
                ReadSeriesName(item),
                ReadCategoryRange(item),
                ReadValuesRange(item),
                ReadCachedCategories(item),
                ReadCachedValues(item),
                ReadSeriesColor(item, context)))
            .ToList();
    }

    private static string? ReadSeriesName(OpenXmlElement series)
    {
        var seriesText = series.Descendants<C.SeriesText>().FirstOrDefault();
        return seriesText?.StringReference?.StringCache?.Elements<C.StringPoint>().FirstOrDefault()?.NumericValue?.Text
            ?? seriesText?.StringReference?.Formula?.Text
            ?? seriesText?.NumericValue?.Text;
    }

    private static string? ReadCategoryRange(OpenXmlElement series)
    {
        var categoryAxisData = series.Descendants<C.CategoryAxisData>().FirstOrDefault();
        return categoryAxisData?.StringReference?.Formula?.Text
            ?? categoryAxisData?.NumberReference?.Formula?.Text
            ?? series.Descendants<C.XValues>().FirstOrDefault()?.NumberReference?.Formula?.Text;
    }

    private static string? ReadValuesRange(OpenXmlElement series)
    {
        return series.Descendants<C.Values>().FirstOrDefault()?.NumberReference?.Formula?.Text
            ?? series.Descendants<C.YValues>().FirstOrDefault()?.NumberReference?.Formula?.Text;
    }

    private static IReadOnlyList<string> ReadCachedCategories(OpenXmlElement series)
    {
        var categoryAxisData = series.Descendants<C.CategoryAxisData>().FirstOrDefault();
        var stringValues = categoryAxisData?.StringReference?.StringCache?.Elements<C.StringPoint>()
            .Select(point => point.NumericValue?.Text)
            .Where(value => value is not null)
            .Select(value => value!)
            .ToList();

        if (stringValues?.Count > 0)
        {
            return stringValues;
        }

        var numericCategories = categoryAxisData?.NumberReference?.NumberingCache?.Elements<C.NumericPoint>()
            .Select(point => point.NumericValue?.Text)
            .Where(value => value is not null)
            .Select(value => value!)
            .ToList();

        if (numericCategories?.Count > 0)
        {
            return numericCategories;
        }

        return series.Descendants<C.XValues>().FirstOrDefault()?.NumberReference?.NumberingCache?.Elements<C.NumericPoint>()
            .Select(point => point.NumericValue?.Text)
            .Where(value => value is not null)
            .Select(value => value!)
            .ToList() ?? [];
    }

    private static IReadOnlyList<double> ReadCachedValues(OpenXmlElement series)
    {
        var values = series.Descendants<C.Values>().FirstOrDefault()?.NumberReference?.NumberingCache?.Elements<C.NumericPoint>()
            .Select(point => point.NumericValue?.Text)
            .Where(value => double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out _))
            .Select(value => double.Parse(value!, CultureInfo.InvariantCulture))
            .ToList();

        if (values?.Count > 0)
        {
            return values;
        }

        return series.Descendants<C.YValues>().FirstOrDefault()?.NumberReference?.NumberingCache?.Elements<C.NumericPoint>()
            .Select(point => point.NumericValue?.Text)
            .Where(value => double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out _))
            .Select(value => double.Parse(value!, CultureInfo.InvariantCulture))
            .ToList() ?? [];
    }

    private static XlsxColorV3? ReadSeriesColor(OpenXmlElement series, DrawingParseContext context)
    {
        var solidFill = series
            .Descendants()
            .FirstOrDefault(element => element.LocalName.Equals("spPr", StringComparison.OrdinalIgnoreCase))
            ?.Descendants()
            .FirstOrDefault(element => element.LocalName.Equals("solidFill", StringComparison.OrdinalIgnoreCase));
        if (solidFill is null)
        {
            return null;
        }

        var rgb = solidFill.ChildElements
            .FirstOrDefault(element => element.LocalName.Equals("srgbClr", StringComparison.OrdinalIgnoreCase))
            ?.GetAttribute("val", string.Empty)
            .Value;
        if (!string.IsNullOrWhiteSpace(rgb))
        {
            var normalized = NormalizeChartRgb(rgb);
            return new XlsxColorV3(normalized, null, null, null, false, normalized);
        }

        var scheme = solidFill.ChildElements
            .FirstOrDefault(element => element.LocalName.Equals("schemeClr", StringComparison.OrdinalIgnoreCase));
        var schemeName = scheme?.GetAttribute("val", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(schemeName) ||
            !context.ThemeColorsByName.TryGetValue(schemeName, out var themeRgb))
        {
            return null;
        }

        var tint = ReadTintFromColorElement(scheme);
        return new XlsxColorV3(null, null, null, tint, false, ApplyTintToChartRgb(themeRgb, tint));
    }

    private static double? ReadTintFromColorElement(OpenXmlElement? colorElement)
    {
        if (colorElement is null)
        {
            return null;
        }

        var lumMod = ReadPercentageTransform(colorElement, "lumMod") ?? 100000d;
        var lumOff = ReadPercentageTransform(colorElement, "lumOff") ?? 0d;

        if (Math.Abs(lumMod - 100000d) < double.Epsilon && Math.Abs(lumOff) < double.Epsilon)
        {
            return null;
        }

        return lumOff > 0 ? lumOff / 100000d : (lumMod / 100000d) - 1d;
    }

    private static double? ReadPercentageTransform(OpenXmlElement element, string localName)
    {
        var value = element.ChildElements
            .FirstOrDefault(child => child.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase))
            ?.GetAttribute("val", string.Empty)
            .Value;

        return double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static string? NormalizeChartRgb(string? rgb)
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

    private static string? ApplyTintToChartRgb(string? rgb, double? tint)
    {
        var normalized = NormalizeChartRgb(rgb);
        if (normalized is null || tint is null || Math.Abs(tint.Value) < double.Epsilon)
        {
            return normalized;
        }

        var channels = Enumerable.Range(0, 3)
            .Select(index => Convert.ToInt32(normalized.Substring(index * 2, 2), 16))
            .Select(channel => ApplyTintToChartChannel(channel, tint.Value).ToString("X2", CultureInfo.InvariantCulture));

        return string.Concat(channels);
    }

    private static int ApplyTintToChartChannel(int channel, double tint)
    {
        var adjusted = tint < 0
            ? channel * (1 + tint)
            : channel * (1 - tint) + 255 * tint;

        return Math.Clamp((int)Math.Round(adjusted), 0, 255);
    }

    private static string ExtensionFromContentType(string contentType)
    {
        return contentType.ToLowerInvariant() switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/jpg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tiff",
            "image/x-emf" => ".emf",
            "image/x-wmf" => ".wmf",
            "image/svg+xml" => ".svg",
            _ => ".bin"
        };
    }

    private static int ParseInt(string? text)
    {
        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;
    }

    private static long ParseLong(string? text)
    {
        return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;
    }

    private sealed class DrawingParseContext(
        string? assetOutputDirectory,
        IReadOnlyDictionary<string, string> themeColorsByName)
    {
        public string? AssetOutputDirectory { get; } = assetOutputDirectory;
        public IReadOnlyDictionary<string, string> ThemeColorsByName { get; } = themeColorsByName;
        public int ChartCount { get; set; }
        public int ImageCount { get; set; }
    }
}
